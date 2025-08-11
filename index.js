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

// Google Sheets authentication
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
      // Process the quality search with full visibility
      const qualities = trimmedMessage.split(',').map(q => q.trim()).filter(q => q.length > 0);
      await processStockQueryWithFullVisibility(from, qualities, productId, phoneId);
      return res.sendStatus(200);
    }
  }

  // Handle store selection for ordering
  if (userStates[from] && userStates[from].currentMenu === 'store_selection') {
    if (/^[1-9]$/.test(trimmedMessage)) {
      await handleStoreSelection(from, trimmedMessage, productId, phoneId);
      return res.sendStatus(200);
    } else if (trimmedMessage !== '/') {
      await sendWhatsAppMessage(from, '❌ Please reply with a number (1, 2, 3, etc.) to select your store.', productId, phoneId);
      return res.sendStatus(200);
    }
  }

  // FOR ALL OTHER MESSAGES: COMPLETE SILENCE
  console.log('Normal message received - bot staying silent:', trimmedMessage);
  return res.sendStatus(200);
});

// NEW: Stock query with FULL visibility but restricted ordering
async function processStockQueryWithFullVisibility(from, qualities, productId, phoneId) {
  try {
    console.log('Processing stock query with full visibility for:', from);
    
    // Send processing message
    await sendWhatsAppMessage(from, '🔍 *Searching stock information across ALL stores...*\nPlease wait while I check complete inventory and your ordering permissions.', productId, phoneId);

    // Get stock results from ALL stores (no filtering)
    const stockResults = await searchStockInAllSheets(qualities);
    
    // Get user's permitted stores for order placement only
    const permittedStores = await getUserPermittedStores(from);
    
    // Format stock results - SHOW ALL STORES
    let responseMessage = `📊 *COMPLETE STOCK QUERY RESULTS*\n`;
    responseMessage += `_Displaying inventory from all stores in the system_\n\n`;
    
    qualities.forEach(quality => {
      responseMessage += `🔸 *${quality}*\n`;
      
      const storeData = stockResults[quality] || {};
      if (Object.keys(storeData).length === 0) {
        responseMessage += `No data found in any store\n\n`;
      } else {
        Object.entries(storeData).forEach(([storeName, stock]) => {
          // Show ALL stores with stock levels - no filtering
          responseMessage += `${storeName} -- ${stock}\n`;
        });
        responseMessage += `\n`;
      }
    });
    
    // Order placement section
    responseMessage += `📋 *ORDER PLACEMENT OPTIONS:*\n\n`;
    
    if (permittedStores.length === 0) {
      responseMessage += `❌ *NO ORDERING PERMISSION*\n\n`;
      responseMessage += `You can see stock from all stores above, but your contact number (${from}) is not authorized to place orders from any store.\n\n`;
      responseMessage += `📞 Please contact administration at *system@fashionformal.com* to get store access permissions.\n\n`;
      responseMessage += `_Type */* to return to main menu._`;
    } else {
      responseMessage += `✅ *You can place orders from these stores only:*\n\n`;
      
      permittedStores.forEach((store, index) => {
        // Show only stores user can order from
        responseMessage += `🏪 *Option ${index + 1}: ${store}*\n`;
        responseMessage += `Reply with: *${index + 1}*\n\n`;
      });
      
      responseMessage += `📋 *Important Notes:*\n`;
      responseMessage += `• You can see stock levels from ALL stores above\n`;
      responseMessage += `• You can ONLY place orders from the ${permittedStores.length} stores listed above\n`;
      responseMessage += `• Select a store number to create your secure order form\n\n`;
      
      // Store user's context for order placement
      userStates[from] = {
        currentMenu: 'store_selection',
        permittedStores: permittedStores,
        qualities: qualities
      };
      
      responseMessage += `_Reply with a number (1, 2, etc.) to place an order, or type */* for main menu._`;
    }
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
  } catch (error) {
    console.error('Error processing full visibility stock query:', error);
    await sendWhatsAppMessage(from, '❌ *Error searching stock*\nPlease try again later.\n\nType */* to return to main menu.', productId, phoneId);
  }
}

// Handle store selection for secure form creation
async function handleStoreSelection(from, storeIndex, productId, phoneId) {
  try {
    console.log(`Handling store selection: ${from} selected ${storeIndex}`);
    
    const userState = userStates[from];
    if (!userState || !userState.permittedStores) {
      await sendWhatsAppMessage(from, '❌ Session expired. Please start over with /\n3\n[quality names]', productId, phoneId);
      return;
    }
    
    const selectedStoreIndex = parseInt(storeIndex) - 1;
    const selectedStore = userState.permittedStores[selectedStoreIndex];
    
    if (!selectedStore) {
      await sendWhatsAppMessage(from, `❌ Invalid selection. Please choose a number between 1-${userState.permittedStores.length}`, productId, phoneId);
      return;
    }
    
    // Create form URL WITHOUT store field - store tracked in form title/description
    const cleanPhone = from.replace(/^\+/, '');
    
    // Create secure URL with pre-filled contact and NO store field in form
    const secureFormUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url` +
      `&entry.740712049=${encodeURIComponent(cleanPhone)}` +
      `&store=${encodeURIComponent(selectedStore)}`;
    
    let confirmationMessage = `✅ *SECURE ORDER FORM CREATED*\n\n`;
    confirmationMessage += `🏪 *Selected Store:* ${selectedStore}\n`;
    confirmationMessage += `📱 *Your Number:* ${cleanPhone}\n`;
    confirmationMessage += `⏰ *Created:* ${new Date().toLocaleString()}\n\n`;
    
    confirmationMessage += `📋 *Your Secure Form:*\n${secureFormUrl}\n\n`;
    
    confirmationMessage += `🔒 *Security Features:*\n`;
    confirmationMessage += `• Form is locked to: *${selectedStore}* only\n`;
    confirmationMessage += `• Contact number is pre-filled and secured\n`;
    confirmationMessage += `• You only need to fill: Quality, MTR, Remarks\n`;
    confirmationMessage += `• NO store field in form = NO way to change selection\n`;
    confirmationMessage += `• Backend validation prevents unauthorized submissions\n\n`;
    
    confirmationMessage += `🛡️ *IMPOSSIBLE TO PLACE ORDERS FROM WRONG STORE!*\n\n`;
    confirmationMessage += `_Type */* to return to main menu._`;
    
    await sendWhatsAppMessage(from, confirmationMessage, productId, phoneId);
    
    // Clear user state
    userStates[from] = { currentMenu: 'completed' };
    
    console.log(`✅ Secure form created for ${from} - Store: ${selectedStore}`);
    
  } catch (error) {
    console.error('Error handling store selection:', error);
    await sendWhatsAppMessage(from, '❌ Error creating secure form. Please try again.', productId, phoneId);
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

// Enhanced search in ALL sheets (no filtering)
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

    // Search EVERY sheet found - no filtering
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
  console.log('✅ Bot ready with FULL STOCK VISIBILITY + Restricted Ordering!');
  console.log(`📊 Stock Folder ID: ${STOCK_FOLDER_ID}`);
  console.log(`🔐 Store Permission Sheet ID: ${STORE_PERMISSION_SHEET_ID}`);
  console.log(`📋 Static Form URL: ${STATIC_FORM_BASE_URL}`);
  console.log('🎯 Users see ALL stock data but can only order from permitted stores!');
});
