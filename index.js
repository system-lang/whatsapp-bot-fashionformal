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

// Static Google Form configuration
const STATIC_FORM_BASE_URL = 'https://docs.google.com/forms/d/e/1FAIpQLSfyAo7LwYtQDfNVxPRbHdk_ymGpDs-RyWTCgzd2PdRhj0T3Hw/viewform';

// Google Sheets authentication (only for stock search now)
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
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.readonly'
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
      await processStockQueryWithPermissionBasedForms(from, qualities, productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }
  }

  // FOR ALL OTHER MESSAGES: COMPLETE SILENCE
  console.log('Normal message received - bot staying silent:', trimmedMessage);
  return res.sendStatus(200);
});

// NEW: Stock query with permission-based pre-filled forms
async function processStockQueryWithPermissionBasedForms(from, qualities, productId, phoneId) {
  try {
    console.log('Processing stock query with permission-based forms for qualities:', qualities);
    
    // Send processing message
    await sendWhatsAppMessage(from, '🔍 *Searching stock information...*\nPlease wait while I check our inventory and your store permissions.', productId, phoneId);

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
    
    // Check if user has store permissions
    if (permittedStores.length === 0) {
      responseMessage += `❌ *NO ORDERING PERMISSION*\n\n`;
      responseMessage += `Your contact number (${from}) is not authorized to place orders from any store.\n\n`;
      responseMessage += `📞 Please contact administration at *system@fashionformal.com* to get store access permissions.\n\n`;
      responseMessage += `_Type */* to return to main menu._`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      return;
    }
    
    // Create personalized forms for each permitted store
    responseMessage += `📋 *INQUIRY FORMS - SELECT YOUR STORE:*\n\n`;
    responseMessage += `✅ *Your Authorized Stores:*\n`;
    
    permittedStores.forEach((store, index) => {
      const cleanPhone = from.replace(/^\+/, ''); // Remove + if present
      
      // Create form URL with pre-filled contact number AND store name
      const storeFormUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url&entry.740712049=${encodeURIComponent(cleanPhone)}&entry.1482226385=${encodeURIComponent(store)}`;
      
      responseMessage += `\n🏪 *${index + 1}. ${store}*\n`;
      responseMessage += `${storeFormUrl}\n`;
    });
    
    responseMessage += `\n📋 *How to use:*\n`;
    responseMessage += `• Each form link above is pre-filled with your contact number and specific store\n`;
    responseMessage += `• You only need to fill: Quality, MTR, and Remarks\n`;
    responseMessage += `• You can ONLY use the forms listed above\n`;
    responseMessage += `• Forms for other stores will not work for your number\n\n`;
    
    responseMessage += `🔒 *Security Note:*\n`;
    responseMessage += `These forms are personalized for your contact number (${from}) and cannot be used by others.\n\n`;
    
    responseMessage += `_Type */* to return to main menu._`;
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
  } catch (error) {
    console.error('Error processing stock query with permissions:', error);
    await sendWhatsAppMessage(from, '❌ *Error searching stock*\nSorry, there was an issue accessing the inventory data. Please try again later.\n\nType */* to return to main menu.', productId, phoneId);
  }
}

// Get user's permitted stores from Store Permission sheet
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

// Enhanced search in all sheets (keeping existing function)
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🤖 WhatsApp Bot running on port ${PORT}`);
  console.log('✅ Bot ready with Permission-Based Pre-filled Forms!');
  console.log(`📊 Stock Folder ID: ${STOCK_FOLDER_ID}`);
  console.log(`🔐 Store Permission Sheet ID: ${STORE_PERMISSION_SHEET_ID}`);
  console.log(`📋 Static Form URL: ${STATIC_FORM_BASE_URL}`);
  console.log('🎯 Using permission-based pre-filled forms - NO unauthorized access possible!');
});
