require('dotenv').config();
const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const app = express();

app.use(express.json());

// Store user states to track which menu they're in
let userStates = {};

// Links for different options
const links = {
  helpTicket: 'https://tinyurl.com/HelpticketFF',
  delegation: 'https://tinyurl.com/DelegationFF',
  leave: 'YOUR_LEAVE_FORM_LINK_HERE'
};

// Your API Token
const MAYTAPI_API_TOKEN = '07d75e68-b94f-485b-9e8c-19e707d176ae';

// Google Sheets configuration
const STOCK_FOLDER_ID = '1QV1cJ9jJZZW2PY24uUY2hefKeUqVHrrf';
const STORE_PERMISSION_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';

// Central response sheet configuration
const CENTRAL_RESPONSE_SHEET_ID = '1nqILVLotV2CSC55bKq0XifyBRm3wAEhg2xKR4V_EcGU';
const RESPONSE_SHEET_NAME = 'submission';

// Store active forms for response collection
let activeForms = new Map(); // formId -> {userPhone, createdAt, processed: Set()}

// FIXED: Google Sheets + Forms authentication with correct scopes
async function getGoogleAuth() {
  try {
    const base64Key = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
    if (!base64Key) {
      throw new Error('GOOGLE_SERVICE_ACCOUNT_KEY environment variable not set');
    }
    
    const keyBuffer = Buffer.from(base64Key, 'base64');
    const keyData = JSON.parse(keyBuffer.toString());
    
    const auth = new google.auth.GoogleAuth({
      credentials: keyData,
      scopes: [
        // Sheets API scopes
        'https://www.googleapis.com/auth/spreadsheets',
        
        // Drive API scopes (required for Forms)
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/drive.resource',
        
        // Forms API scopes (the required ones!)
        'https://www.googleapis.com/auth/forms.body',
        'https://www.googleapis.com/auth/forms'
      ],
    });
    
    return auth;
  } catch (error) {
    console.error('Error setting up Google auth:', error);
    throw error;
  }
}

app.post('/webhook', async (req, res) => {
  console.log('Full webhook data:', JSON.stringify(req.body, null, 2));

  const message = req.body.message?.text;
  const from = req.body.user?.phone;
  const productId = req.body.product_id || req.body.productId;
  const phoneId = req.body.phone_id || req.body.phoneId;

  console.log('Extracted from webhook:');
  console.log('Message:', message);
  console.log('From:', from);

  // IMMEDIATELY EXIT if message is empty, null, or invalid
  if (!message || typeof message !== 'string') {
    console.log('No valid message - staying silent');
    return res.sendStatus(200);
  }

  const trimmedMessage = message.trim();

  if (trimmedMessage === '') {
    console.log('Empty message after trim - staying silent');
    return res.sendStatus(200);
  }

  // Handle main menu trigger - ONLY for "/"
  if (trimmedMessage === '/') {
    console.log('Menu command received - activating bot');
    userStates[from] = { currentMenu: 'main' };
    const mainMenu = `🏠 *MAIN MENU*
Please select an option:

1️⃣ Ticket
2️⃣ Order Query  
3️⃣ Stock Query
4️⃣ Document

_Type the number to continue..._`;
    
    await sendWhatsAppMessage(from, mainMenu, productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle menu selections ONLY if user previously used "/"
  if (userStates[from] && userStates[from].currentMenu === 'main') {
    if (trimmedMessage === '1') {
      const ticketMenu = `🎫 *TICKET OPTIONS*
Click the links below to access forms directly:

🆘 *HELP TICKET*
${links.helpTicket}

🏖️ *LEAVE FORM*
${links.leave}

👥 *DELEGATION*
${links.delegation}

_Type */* to return to main menu._`;
      
      await sendWhatsAppMessage(from, ticketMenu, productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2') {
      await sendWhatsAppMessage(from, '🔍 *ORDER QUERY*\nThis feature is coming soon!\n\nType */* to return to main menu.', productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
      // Stock Query - Ask for qualities
      userStates[from].currentMenu = 'stock_query';
      const stockQueryPrompt = `📊 *STOCK QUERY*
Please enter the Quality names you want to search for.

*Multiple qualities:* Separate with commas
*Example:* LTS8156, ETCH8029, Quality3

_Type your quality names below:_`;
      
      await sendWhatsAppMessage(from, stockQueryPrompt, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '4') {
      await sendWhatsAppMessage(from, '📄 *DOCUMENT*\nThis feature is coming soon!\n\nType */* to return to main menu.', productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }

    // Invalid menu option
    await sendWhatsAppMessage(from, '❌ Invalid option. Please select 1, 2, 3, or 4.\n\nType */* to see the main menu again.', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle stock query input
  if (userStates[from] && userStates[from].currentMenu === 'stock_query') {
    if (trimmedMessage !== '/') {
      // Process the quality search
      const qualities = trimmedMessage.split(',').map(q => q.trim()).filter(q => q.length > 0);
      await processEnhancedStockQuery(from, qualities, productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }
  }

  // FOR ALL OTHER MESSAGES: COMPLETE SILENCE
  console.log('Normal message received - bot staying silent:', trimmedMessage);
  return res.sendStatus(200);
});

// ENHANCED: Stock query with dynamic dropdown forms and central responses
async function processEnhancedStockQuery(from, qualities, productId, phoneId) {
  try {
    console.log('Processing enhanced stock query for qualities:', qualities);
    
    // Send processing message
    await sendWhatsAppMessage(from, '🔍 *Searching stock information...*\nPlease wait while I check our inventory.', productId, phoneId);

    // Get stock results
    const stockResults = await searchStockInAllSheets(qualities);
    
    // Get user's permitted stores
    const permittedStores = await getUserPermittedStores(from);
    
    // Format stock results
    let responseMessage = `📊 *STOCK QUERY RESULTS*\n\n`;
    
    qualities.forEach(quality => {
      responseMessage += `🔸 *${quality}*\n`;
      
      const storeData = stockResults[quality] || {};
      if (Object.keys(storeData).length === 0) {
        responseMessage += `No data found in any store\n\n`;
      } else {
        Object.entries(storeData).forEach(([storeName, stock]) => {
          responseMessage += `${storeName} -- ${stock}\n`;
        });
        responseMessage += `\n`;
      }
    });
    
    // Create dynamic form with REAL dropdowns if user has store permissions
    if (permittedStores.length > 0) {
      console.log(`Creating dynamic dropdown form for ${from} with stores:`, permittedStores);
      const formUrl = await createDynamicFormWithCentralResponse(qualities, permittedStores, from);
      
      if (formUrl && formUrl !== 'Form creation temporarily unavailable') {
        responseMessage += `📋 *INQUIRY FORM*\nTo place an inquiry for any of these qualities:\n${formUrl}\n\n`;
        responseMessage += `*Features:*\n`;
        responseMessage += `• Store dropdown: Only your permitted stores\n`;
        responseMessage += `• Quality dropdown: Only searched items\n`;
        responseMessage += `• All responses saved centrally\n\n`;
      } else {
        responseMessage += `📋 *INQUIRY FORM - TEMPORARILY UNAVAILABLE*\n`;
        responseMessage += `Google Forms API is experiencing issues. Please contact us directly:\n\n`;
        responseMessage += `📞 *Contact:* ${from}\n`;
        responseMessage += `🏪 *Your Stores:* ${permittedStores.join(', ')}\n`;
        responseMessage += `📦 *Qualities:* ${qualities.join(', ')}\n\n`;
        responseMessage += `We'll create your inquiry manually and get back to you soon!\n\n`;
      }
    } else {
      console.log(`No store permissions found for ${from}`);
    }
    
    responseMessage += `_Type */* to return to main menu._`;
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
  } catch (error) {
    console.error('Error processing enhanced stock query:', error);
    await sendWhatsAppMessage(from, '❌ *Error searching stock*\nSorry, there was an issue accessing the inventory data. Please try again later.\n\nType */* to return to main menu.', productId, phoneId);
  }
}

// ENHANCED: Create dynamic form with comprehensive debug logging
async function createDynamicFormWithCentralResponse(qualities, permittedStores, userPhone) {
  console.log('\n=== 🚀 GOOGLE FORMS API DEBUG SESSION ===');
  console.log(`📱 User Phone: ${userPhone}`);
  console.log(`🏪 Permitted Stores: [${permittedStores.join(', ')}]`);
  console.log(`📦 Searched Qualities: [${qualities.join(', ')}]`);
  console.log('⏰ Timestamp:', new Date().toISOString());
  
  try {
    // STEP 1: Authentication Test
    console.log('\n--- STEP 1: Testing Authentication ---');
    const auth = await getGoogleAuth();
    console.log('✅ Google Auth object created successfully');
    
    const authClient = await auth.getClient();
    console.log('✅ Auth client obtained successfully');
    console.log('🔑 Auth client type:', authClient.constructor.name);
    
    // STEP 2: Forms API Client Test
    console.log('\n--- STEP 2: Testing Forms API Client ---');
    const forms = google.forms({ version: 'v1', auth: authClient });
    console.log('✅ Google Forms API client created successfully');
    console.log('🔗 Forms API endpoint:', 'https://forms.googleapis.com/v1/forms');
    
    // STEP 3: Pre-flight Check
    console.log('\n--- STEP 3: Pre-flight API Check ---');
    console.log('📋 Preparing form creation request...');
    console.log('📝 Form title:', `Stock Inquiry - ${userPhone}`);
    
    const requestBody = {
      info: {
        title: `Stock Inquiry - ${userPhone}`
      }
    };
    console.log('📤 Request body:', JSON.stringify(requestBody, null, 2));
    
    // STEP 4: Attempt Form Creation
    console.log('\n--- STEP 4: Creating Google Form ---');
    console.log('⚡ Sending CREATE request to Google Forms API...');
    
    const startTime = Date.now();
    
    const form = await forms.forms.create({
      requestBody: requestBody
    });
    
    const endTime = Date.now();
    const responseTime = endTime - startTime;
    
    console.log(`✅ Form creation SUCCESS! (${responseTime}ms)`);
    console.log('🆔 Form ID:', form.data.formId);
    console.log('📊 Form URL:', `https://docs.google.com/forms/d/${form.data.formId}/viewform`);
    console.log('📋 Full form response:', JSON.stringify(form.data, null, 2));

    const formId = form.data.formId;

    // STEP 5: Building Batch Update Requests
    console.log('\n--- STEP 5: Preparing Batch Updates ---');
    
    const requests = [];
    
    // Description update
    console.log('📝 Adding description update...');
    requests.push({
      updateFormInfo: {
        info: {
          title: `Stock Inquiry - ${userPhone}`,
          description: `Submit your inquiry for selected qualities.\n\nUser: ${userPhone}\nGenerated: ${new Date().toLocaleString()}`
        },
        updateMask: 'description'
      }
    });
    
    // Store Name dropdown
    console.log(`🏪 Adding Store Name dropdown with ${permittedStores.length} options...`);
    requests.push({
      createItem: {
        item: {
          title: 'Store Name',
          description: 'Select the store you want to inquire about',
          questionItem: {
            question: {
              required: true,
              choiceQuestion: {
                type: 'DROP_DOWN',
                options: permittedStores.map(store => ({ value: store }))
              }
            }
          }
        },
        location: { index: 0 }
      }
    });

    // Quality dropdown
    console.log(`📦 Adding Quality dropdown with ${qualities.length} options...`);
    requests.push({
      createItem: {
        item: {
          title: 'Quality',
          description: 'Select the quality you want to inquire about',
          questionItem: {
            question: {
              required: true,
              choiceQuestion: {
                type: 'DROP_DOWN',
                options: qualities.map(quality => ({ value: quality }))
              }
            }
          }
        },
        location: { index: 1 }
      }
    });

    // MTR field
    console.log('📏 Adding MTR number field...');
    requests.push({
      createItem: {
        item: {
          title: 'MTR (Meters Required)',
          description: 'Enter the quantity you need in meters',
          questionItem: {
            question: {
              required: true,
              textQuestion: {
                paragraph: false
              }
            }
          }
        },
        location: { index: 2 }
      }
    });

    // Remarks field
    console.log('💬 Adding Remarks text area...');
    requests.push({
      createItem: {
        item: {
          title: 'Remarks',
          description: 'Any additional notes or requirements (optional)',
          questionItem: {
            question: {
              required: false,
              textQuestion: {
                paragraph: true
              }
            }
          }
        },
        location: { index: 3 }
      }
    });

    console.log(`📋 Total batch requests prepared: ${requests.length}`);
    
    // STEP 6: Apply Batch Updates
    console.log('\n--- STEP 6: Applying Batch Updates ---');
    console.log('⚡ Sending BATCH UPDATE request...');
    
    const batchStartTime = Date.now();
    
    await forms.forms.batchUpdate({
      formId: formId,
      requestBody: {
        requests: requests
      }
    });
    
    const batchEndTime = Date.now();
    const batchResponseTime = batchEndTime - batchStartTime;
    
    console.log(`✅ Batch update SUCCESS! (${batchResponseTime}ms)`);
    
    // STEP 7: Configure Form Settings
    console.log('\n--- STEP 7: Configuring Form Settings ---');
    console.log('⚙️ Setting submit button and confirmation message...');
    
    const settingsStartTime = Date.now();
    
    await forms.forms.batchUpdate({
      formId: formId,
      requestBody: {
        requests: [
          {
            updateSettings: {
              settings: {
                submitButtonText: 'Submit Inquiry',
                confirmationMessage: 'Thank you! Your inquiry has been submitted successfully.'
              },
              updateMask: 'submitButtonText,confirmationMessage'
            }
          }
        ]
      }
    });
    
    const settingsEndTime = Date.now();
    const settingsResponseTime = settingsEndTime - settingsStartTime;
    
    console.log(`✅ Settings update SUCCESS! (${settingsResponseTime}ms)`);

    // STEP 8: Track Form for Response Collection
    console.log('\n--- STEP 8: Registering Form for Response Collection ---');
    
    activeForms.set(formId, {
      userPhone: userPhone,
      createdAt: new Date(),
      processed: new Set()
    });
    
    console.log(`✅ Form ${formId} registered for response collection`);
    console.log(`📊 Total active forms: ${activeForms.size}`);

    // STEP 9: Final Success Summary
    console.log('\n--- 🎉 FORM CREATION COMPLETE SUCCESS! ---');
    const totalTime = Date.now() - startTime;
    console.log(`⏱️  Total time: ${totalTime}ms`);
    console.log(`📋 Form ID: ${formId}`);
    console.log(`🔗 Form URL: https://docs.google.com/forms/d/${formId}/viewform`);
    console.log('🏪 Store dropdown options:', permittedStores);
    console.log('📦 Quality dropdown options:', qualities);
    console.log('=== 🚀 DEBUG SESSION COMPLETE ===\n');
    
    return `https://docs.google.com/forms/d/${formId}/viewform`;
    
  } catch (error) {
    // COMPREHENSIVE ERROR LOGGING
    console.log('\n--- ❌ ERROR OCCURRED ---');
    console.log('🔴 Error type:', error.constructor.name);
    console.log('🔴 Error message:', error.message);
    console.log('🔴 Error status:', error.status || 'unknown');
    console.log('🔴 Error code:', error.code || 'unknown');
    
    if (error.response) {
      console.log('📤 Request details:');
      console.log('   • URL:', error.response.config?.url);
      console.log('   • Method:', error.response.config?.method);
      console.log('   • Body:', error.response.config?.body);
      
      console.log('📥 Response details:');
      console.log('   • Status:', error.response.status);
      console.log('   • Status Text:', error.response.statusText);
      console.log('   • Headers:', JSON.stringify(error.response.headers, null, 2));
      
      if (error.response.data) {
        console.log('   • Response Data:', JSON.stringify(error.response.data, null, 2));
      }
    }
    
    if (error.errors) {
      console.log('🔴 Detailed errors:');
      error.errors.forEach((err, index) => {
        console.log(`   ${index + 1}. ${err.message} (${err.reason})`);
      });
    }
    
    console.log('🔴 Full error object:', JSON.stringify(error, null, 2));
    console.log('=== ❌ ERROR DEBUG SESSION COMPLETE ===\n');
    
    return 'Form creation temporarily unavailable';
  }
}

// RESPONSE COLLECTION: Get user's permitted stores from Store Permission sheet
async function getUserPermittedStores(phoneNumber) {
  try {
    console.log(`Getting permitted stores for phone: ${phoneNumber}`);
    
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    // Read Store Permission sheet
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: STORE_PERMISSION_SHEET_ID,
      range: 'A:B',
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log('No data found in Store Permission sheet');
      return [];
    }

    const permittedStores = [];
    
    // Clean phone number for comparison (remove country code variations)
    const cleanPhone = phoneNumber.replace(/^\+91|^91|^0/, '');
    console.log(`Cleaned phone number: ${cleanPhone}`);
    
    for (let i = 1; i < rows.length; i++) { // Skip header
      const row = rows[i];
      if (!row || !row[0] || !row[1]) continue;
      
      const contactNumber = row[0].toString().replace(/^\+91|^91|^0/, '');
      const storeName = row[1].toString().trim();
      
      console.log(`Checking: ${contactNumber} === ${cleanPhone} for store: ${storeName}`);
      
      if (contactNumber === cleanPhone) {
        permittedStores.push(storeName);
        console.log(`Added permitted store: ${storeName}`);
      }
    }
    
    console.log(`Found ${permittedStores.length} permitted stores for ${phoneNumber}:`, permittedStores);
    return permittedStores;
    
  } catch (error) {
    console.error('Error getting permitted stores:', error);
    return [];
  }
}

// SEARCH: Enhanced search in all sheets
async function searchStockInAllSheets(qualities) {
  const results = {};
  
  // Initialize results structure
  qualities.forEach(quality => {
    results[quality] = {};
  });

  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    // Find ALL spreadsheets in the folder
    const folderFiles = await drive.files.list({
      q: `'${STOCK_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    console.log('Found ALL files in folder:', folderFiles.data.files);

    // Search EVERY sheet found
    for (const file of folderFiles.data.files) {
      console.log(`Searching in store: ${file.name} (${file.id})`);

      try {
        // Get data from columns A and E
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: file.id,
          range: 'A:E',
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
          console.log(`No data found in ${file.name}`);
          continue;
        }

        console.log(`Found ${rows.length} rows in ${file.name}`);
        
        // Search for each quality with multiple matching strategies
        qualities.forEach(searchQuality => {
          const qualityUpper = searchQuality.toUpperCase().trim();
          const qualityLower = searchQuality.toLowerCase().trim();
          const qualityOriginal = searchQuality.trim();
          
          console.log(`Searching for quality: "${searchQuality}" in ${file.name}`);
          
          for (let i = 1; i < rows.length; i++) { // Skip header row
            const row = rows[i];
            if (!row || !row[0]) continue;
            
            const cellQuality = row[0].toString().trim();
            const cellQualityUpper = cellQuality.toUpperCase();
            const cellQualityLower = cellQuality.toLowerCase();
            
            // Multiple matching strategies
            if (cellQuality === qualityOriginal || 
                cellQualityUpper === qualityUpper || 
                cellQualityLower === qualityLower ||
                cellQuality.includes(qualityOriginal) ||
                cellQualityUpper.includes(qualityUpper) ||
                cellQualityLower.includes(qualityLower)) {
              
              const stockValue = row[4] ? row[4].toString().trim() : '0';
              console.log(`FOUND MATCH! ${searchQuality} in ${file.name}: ${stockValue}`);
              results[searchQuality][file.name] = stockValue;
              break;
            }
          }
          
          // Log if not found
          if (!results[searchQuality][file.name]) {
            console.log(`Quality "${searchQuality}" NOT found in ${file.name}`);
          }
        });

      } catch (sheetError) {
        console.error(`Error accessing sheet ${file.name}:`, sheetError.message);
      }
    }

    console.log('Final search results:', JSON.stringify(results, null, 2));
    return results;

  } catch (error) {
    console.error('Error searching sheets:', error);
    throw error;
  }
}

// COLLECTION: Collect responses from all active forms
async function collectFormResponses() {
  try {
    console.log(`📥 Starting response collection for ${activeForms.size} active forms...`);
    
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const forms = google.forms({ version: 'v1', auth: authClient });
    
    let totalCollected = 0;
    
    for (const [formId, formInfo] of activeForms) {
      try {
        console.log(`Checking form ${formId} for user ${formInfo.userPhone}`);
        
        // Get responses from this form
        const responses = await forms.forms.responses.list({
          formId: formId
        });

        if (responses.data.responses && responses.data.responses.length > 0) {
          console.log(`Found ${responses.data.responses.length} responses in form ${formId}`);
          
          for (const response of responses.data.responses) {
            // Check if this response was already processed
            if (!formInfo.processed.has(response.responseId)) {
              await forwardResponseToCentralSheet(response, formInfo.userPhone, formId);
              formInfo.processed.add(response.responseId);
              totalCollected++;
              console.log(`✅ Processed response ${response.responseId}`);
            } else {
              console.log(`⏭️  Skipped already processed response ${response.responseId}`);
            }
          }
        } else {
          console.log(`No responses found in form ${formId}`);
        }
        
      } catch (formError) {
        console.error(`Error processing form ${formId}:`, formError.message);
      }
    }
    
    console.log(`📥 Response collection complete. Collected ${totalCollected} new responses.`);
    
  } catch (error) {
    console.error('Error in response collection:', error);
  }
}

// FORWARD: Forward response to central submission sheet
async function forwardResponseToCentralSheet(response, userPhone, formId) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    // Extract answers from response
    const answers = response.answers || {};
    const answersList = Object.values(answers);
    
    // Extract values (assuming order: Store Name, Quality, MTR, Remarks)
    const storeName = answersList[0]?.textAnswers?.answers?.[0]?.value || '';
    const quality = answersList[1]?.textAnswers?.answers?.[0]?.value || '';
    const mtr = answersList[2]?.textAnswers?.answers?.[0]?.value || '';
    const remarks = answersList[3]?.textAnswers?.answers?.[0]?.value || '';

    const values = [
      new Date().toISOString(),  // A: Timestamp
      userPhone,                 // B: Contact Number
      storeName,                 // C: Store Name
      quality,                   // D: Quality
      mtr,                       // E: MTR
      remarks,                   // F: Remarks
      formId,                    // G: Form ID
      response.responseId        // H: Response ID (for duplicate prevention)
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: CENTRAL_RESPONSE_SHEET_ID,
      range: `${RESPONSE_SHEET_NAME}!A:H`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [values]
      }
    });

    console.log(`✅ Forwarded response ${response.responseId} to central submission sheet`);
    
  } catch (error) {
    console.error(`Error forwarding response to central sheet:`, error);
  }
}

// WhatsApp message sending
async function sendWhatsAppMessage(to, message, productId, phoneId) {
  try {
    console.log('Sending API request with WEBHOOK DATA:');
    console.log('Product ID:', productId);
    console.log('Phone ID:', phoneId);
    console.log('To:', to);

    const response = await axios.post(
      `https://api.maytapi.com/api/${productId}/${phoneId}/sendMessage`,
      {
        to_number: to,
        type: "text",
        message: message
      },
      {
        headers: {
          'x-maytapi-key': MAYTAPI_API_TOKEN,
          'Content-Type': 'application/json'
        }
      }
    );
    console.log('Message sent successfully:', response.data);
  } catch (error) {
    console.error('Error sending message:', error.response?.data || error.message);
  }
}

// Start response collection interval (every 5 minutes)
setInterval(collectFormResponses, 5 * 60 * 1000);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🤖 WhatsApp Bot running on port ${PORT}`);
  console.log('✅ Bot ready with Dynamic Dropdown Forms + Central Response Collection!');
  console.log(`📊 Stock Folder ID: ${STOCK_FOLDER_ID}`);
  console.log(`🔐 Store Permission Sheet ID: ${STORE_PERMISSION_SHEET_ID}`);
  console.log(`📋 Central Response Sheet ID: ${CENTRAL_RESPONSE_SHEET_ID}`);
  console.log(`📝 Response Sheet Name: ${RESPONSE_SHEET_NAME}`);
  console.log('🔄 Response collection will run every 5 minutes');
  
  // Initial response collection after 1 minute
  setTimeout(collectFormResponses, 60 * 1000);
});
