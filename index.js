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

// Order Query Configuration
const LIVE_SHEET_ID = '1AxjCHsMxYUmEULaW1LxkW78g0Bv9fp4PkZteJO82uEA';
const LIVE_SHEET_NAME = 'FMS';
const COMPLETED_ORDER_FOLDER_ID = '1kgdPdnUK-FsnKZDE5yW6vtRf2H9d3YRE';

// Production stages configuration - FIXED to check all stages
const PRODUCTION_STAGES = [
  { name: 'CUT', column: 'O', nextStage: 'FUS' },
  { name: 'FUS', column: 'U', nextStage: 'PAS' },
  { name: 'PAS', column: 'AA', nextStage: 'MAK' },
  { name: 'MAK', column: 'AG', nextStage: 'BH' },
  { name: 'BH', column: 'AS', nextStage: 'BS' },
  { name: 'BS', column: 'AY', nextStage: 'QC' },
  { name: 'QC', column: 'BE', nextStage: 'ALT' },
  { name: 'ALT', column: 'BK', nextStage: 'IRO' },
  { name: 'IRO', column: 'BT', nextStage: 'Dispatch (Factory)' },
  { name: 'Dispatch (Factory)', column: 'BZ', nextStage: 'Dispatch (HO)' },
  { name: 'Dispatch (HO)', column: 'CG', nextStage: 'COMPLETED', dispatchDateColumn: 'CH' }
];

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

  if (!message || typeof message !== 'string') {
    console.log('No valid message - staying silent');
    return res.sendStatus(200);
  }

  const trimmedMessage = message.trim();

  if (trimmedMessage === '') {
    console.log('Empty message after trim - staying silent');
    return res.sendStatus(200);
  }

  // Handle main menu trigger
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

  // Handle menu selections
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
      userStates[from].currentMenu = 'order_query';
      const orderQueryMenu = `📦 *ORDER QUERY*
Please select the product category:

1️⃣ Shirting
2️⃣ Jacket  
3️⃣ Trouser

_Type the number to continue..._`;
      
      await sendWhatsAppMessage(from, orderQueryMenu, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
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

    await sendWhatsAppMessage(from, '❌ Invalid option. Please select 1, 2, 3, or 4.\n\nType */* to see the main menu again.', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle order query category selection
  if (userStates[from] && userStates[from].currentMenu === 'order_query') {
    if (trimmedMessage === '1') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Shirting' };
      await sendWhatsAppMessage(from, `👔 *SHIRTING ORDER QUERY*\n\nPlease enter your Order Number(s):\n\n*Single order:* ABC123\n*Multiple orders:* ABC123, DEF456, GHI789\n\n_Type your order numbers below:_`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Jacket' };
      await sendWhatsAppMessage(from, `🧥 *JACKET ORDER QUERY*\n\nPlease enter your Order Number(s):\n\n*Single order:* ABC123\n*Multiple orders:* ABC123, DEF456, GHI789\n\n_Type your order numbers below:_`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Trouser' };
      await sendWhatsAppMessage(from, `👖 *TROUSER ORDER QUERY*\n\nPlease enter your Order Number(s):\n\n*Single order:* ABC123\n*Multiple orders:* ABC123, DEF456, GHI789\n\n_Type your order numbers below:_`, productId, phoneId);
      return res.sendStatus(200);
    }

    await sendWhatsAppMessage(from, '❌ Invalid option. Please select 1, 2, or 3.\n\nType */* to return to main menu.', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle order number input
  if (userStates[from] && userStates[from].currentMenu === 'order_number_input') {
    if (trimmedMessage !== '/') {
      const category = userStates[from].category;
      const orderNumbers = trimmedMessage.split(',').map(order => order.trim()).filter(order => order.length > 0);
      
      await processOrderQuery(from, category, orderNumbers, productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }
  }

  // Handle stock query input
  if (userStates[from] && userStates[from].currentMenu === 'stock_query') {
    if (trimmedMessage !== '/') {
      const qualities = trimmedMessage.split(',').map(q => q.trim()).filter(q => q.length > 0);
      await processStockQueryWithMultipleOrders(from, qualities, productId, phoneId);
      return res.sendStatus(200);
    }
  }

  // Handle multiple order selection
  if (userStates[from] && userStates[from].currentMenu === 'multiple_order_selection') {
    if (trimmedMessage !== '/') {
      await handleMultipleOrderSelectionWithHiddenField(from, trimmedMessage, productId, phoneId);
      return res.sendStatus(200);
    }
  }

  console.log('Normal message received - bot staying silent:', trimmedMessage);
  return res.sendStatus(200);
});

// Process order query
async function processOrderQuery(from, category, orderNumbers, productId, phoneId) {
  try {
    console.log(`Processing order query for ${category}: ${orderNumbers.join(', ')}`);
    
    await sendWhatsAppMessage(from, `🔍 *Checking ${category} orders...*\nPlease wait while I search for your order status.`, productId, phoneId);

    let responseMessage = `📦 *${category.toUpperCase()} ORDER STATUS*\n\n`;
    
    for (const orderNum of orderNumbers) {
      console.log(`Searching for order: ${orderNum}`);
      
      const orderStatus = await searchOrderStatus(orderNum, category);
      
      responseMessage += `🔸 *Order: ${orderNum}*\n`;
      responseMessage += `${orderStatus.message}\n\n`;
    }
    
    responseMessage += `_Type */* to return to main menu._`;
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
  } catch (error) {
    console.error('Error processing order query:', error);
    await sendWhatsAppMessage(from, '❌ *Error checking orders*\nPlease try again later.\n\nType */* to return to main menu.', productId, phoneId);
  }
}

// Search order status
async function searchOrderStatus(orderNumber, category) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    // Step 1: Search in live sheet
    console.log(`Searching ${orderNumber} in live sheet (FMS)`);
    
    const liveSheetResult = await searchInLiveSheet(sheets, orderNumber);
    if (liveSheetResult.found) {
      console.log(`Order ${orderNumber} found in live sheet`);
      return liveSheetResult;
    }

    // Step 2: Search in completed order folder
    console.log(`Order not found in live sheet, searching completed order folder`);
    
    const folderFiles = await drive.files.list({
      q: `'${COMPLETED_ORDER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    console.log(`Found ${folderFiles.data.files.length} completed order sheets`);

    for (const file of folderFiles.data.files) {
      console.log(`Searching in ${file.name}`);
      
      try {
        const completedResult = await searchInCompletedSheetSimplified(sheets, file.id, orderNumber);
        if (completedResult.found) {
          console.log(`Order ${orderNumber} found in ${file.name}`);
          return completedResult;
        }
      } catch (error) {
        console.log(`Error searching ${file.name}: ${error.message}`);
        continue;
      }
    }

    console.log(`Order ${orderNumber} not found in any system`);
    return { 
      found: false, 
      message: '❌ Order not found in system. Please contact responsible person.\n\nThank you.' 
    };

  } catch (error) {
    console.error('Error in searchOrderStatus:', error);
    return { 
      found: false, 
      message: '❌ Error occurred while searching order. Please contact responsible person.\n\nThank you.' 
    };
  }
}

// Search in live sheet
async function searchInLiveSheet(sheets, orderNumber) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: LIVE_SHEET_ID,
      range: `${LIVE_SHEET_NAME}!A:CH`,
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return { found: false };
    }

    // Find order in column D (index 3)
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[3]) continue;
      
      if (row[1].toString().trim() === orderNumber.trim()) {
        console.log(`Order ${orderNumber} found in FMS sheet at row ${i + 1}`);
        
        // FIXED: Check all production stages
        const stageStatus = checkProductionStages(row);
        return {
          found: true,
          message: stageStatus.message,
          location: 'Live Sheet (FMS)'
        };
      }
    }

    return { found: false };

  } catch (error) {
    console.error('Error searching live sheet:', error);
    return { found: false };
  }
}

// Search in completed order sheets
async function searchInCompletedSheetSimplified(sheets, sheetId, orderNumber) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'A:CH',
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return { found: false };
    }

    // Find order in column D
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[3]) continue;
      
      if (row[1].toString().trim() === orderNumber.trim()) {
        console.log(`Order ${orderNumber} found in completed sheet at row ${i + 1}`);
        
        // Get dispatch date from column CH (index 87)
        const dispatchDate = row ? row.toString().trim() : 'Date not available';
        
        return {
          found: true,
          message: `✅ Order got dispatched on ${dispatchDate}`,
          location: 'Completed Orders'
        };
      }
    }

    return { found: false };

  } catch (error) {
    console.error('Error searching completed sheet:', error);
    return { found: false };
  }
}

// FIXED: Check ALL production stages for last completed one
function checkProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let lastCompletedStageIndex = -1;
    let hasAnyStage = false;

    console.log('🔍 Checking ALL production stages for last completed...');

    // Check EVERY production stage (don't stop at first empty)
    for (let i = 0; i < PRODUCTION_STAGES.length; i++) {
      const stage = PRODUCTION_STAGES[i];
      const columnIndex = columnToIndex(stage.column);
      const cellValue = row[columnIndex] ? row[columnIndex].toString().trim() : '';
      
      console.log(`Stage ${stage.name} (Col ${stage.column}): "${cellValue}"`);
      
      if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
        lastCompletedStage = stage;
        lastCompletedStageIndex = i;
        hasAnyStage = true;
        console.log(`✅ Stage ${stage.name} completed: ${cellValue} (Index: ${i})`);
      } else {
        console.log(`❌ Stage ${stage.name} not completed`);
      }
    }

    console.log(`🎯 Last completed stage: ${lastCompletedStage ? lastCompletedStage.name : 'None'} (Index: ${lastCompletedStageIndex})`);

    // Generate status message based on findings
    if (!hasAnyStage) {
      return { message: '🟡 Order is currently under process' };
    }

    // Check if it's the final stage (Dispatch HO)
    if (lastCompletedStage && lastCompletedStage.name === 'Dispatch (HO)') {
      // Get dispatch date from column CH
      const dispatchDateIndex = columnToIndex(lastCompletedStage.dispatchDateColumn);
      const dispatchDate = row[dispatchDateIndex] ? row[dispatchDateIndex].toString().trim() : 'Date not available';
      return { message: `✅ Order has been dispatched from HO on ${dispatchDate}` };
    }

    // For any other completed stage, show current stage and next stage
    if (lastCompletedStage) {
      return { 
        message: `🔄 Order is currently completed ${lastCompletedStage.name} stage and processed to ${lastCompletedStage.nextStage} stage` 
      };
    }

    return { message: '❌ Error determining order status' };

  } catch (error) {
    console.error('❌ Error checking production stages:', error);
    return { message: '❌ Error checking order status' };
  }
}

// Convert column letter to index
function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

// STOCK QUERY FUNCTIONS - COMPLETE IMPLEMENTATION
async function processStockQueryWithMultipleOrders(from, qualities, productId, phoneId) {
  try {
    console.log('Processing stock query with multiple orders for:', from);
    
    await sendWhatsAppMessage(from, '🔍 *Searching stock information...*\nPlease wait.', productId, phoneId);

    const stockResults = await searchStockInAllSheets(qualities);
    const permittedStores = await getUserPermittedStores(from);
    
    let responseMessage = `📊 *STOCK QUERY RESULTS*\n\n`;
    
    qualities.forEach(quality => {
      responseMessage += `🔸 *${quality}*\n`;
      
      const storeData = stockResults[quality] || {};
      if (Object.keys(storeData).length === 0) {
        responseMessage += `No data found\n\n`;
      } else {
        Object.entries(storeData).forEach(([storeName, stock]) => {
          responseMessage += `${storeName} -- ${stock}\n`;
        });
        responseMessage += `\n`;
      }
    });
    
    if (permittedStores.length === 0) {
      responseMessage += `❌ *NO ORDERING PERMISSION*\n\n`;
      responseMessage += `Contact: *system@fashionformal.com*\n\n`;
      responseMessage += `_Type */* for main menu._`;
    } else {
      responseMessage += `📋 *PLACE MULTIPLE ORDERS:*\n\n`;
      responseMessage += `*Format:* Quality-StoreNumber, Quality-StoreNumber\n`;
      responseMessage += `*Example:* LTS8156-1, ETCH8029-2\n\n`;
      
      responseMessage += `*Your Store Numbers:*\n`;
      permittedStores.forEach((store, index) => {
        responseMessage += `${index + 1}. ${store}\n`;
      });
      
      responseMessage += `\nReply with combinations or single store number for all items.`;
      
      userStates[from] = {
        currentMenu: 'multiple_order_selection',
        permittedStores: permittedStores,
        qualities: qualities
      };
    }
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
  } catch (error) {
    console.error('Error:', error);
    await sendWhatsAppMessage(from, '❌ *Error searching stock*\n\nType */* to return to main menu.', productId, phoneId);
  }
}

async function getUserPermittedStores(phoneNumber) {
  try {
    console.log(`Getting permitted stores for phone: ${phoneNumber}`);
    
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: STORE_PERMISSION_SHEET_ID,
      range: 'A:B',
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return [];
    }

    const permittedStores = [];
    const cleanPhone = phoneNumber.replace(/^\+91|^91|^0/, '');
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[0] || !row[2]) continue;
      
      const contactNumber = row.toString().replace(/^\+91|^91|^0/, '');
      const storeName = row[2].toString().trim();
      
      if (contactNumber === cleanPhone) {
        permittedStores.push(storeName);
      }
    }
    
    return permittedStores;
    
  } catch (error) {
    console.error('Error getting permitted stores:', error);
    return [];
  }
}

async function searchStockInAllSheets(qualities) {
  const results = {};
  
  qualities.forEach(quality => {
    results[quality] = {};
  });

  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    const folderFiles = await drive.files.list({
      q: `'${STOCK_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    console.log('Found stock files:', folderFiles.data.files.length);

    for (const file of folderFiles.data.files) {
      console.log(`Searching in stock sheet: ${file.name}`);
      
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: file.id,
          range: 'A:E',
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
          console.log(`No data in ${file.name}`);
          continue;
        }
        
        qualities.forEach(searchQuality => {
          const qualityUpper = searchQuality.toUpperCase().trim();
          const qualityLower = searchQuality.toLowerCase().trim();
          const qualityOriginal = searchQuality.trim();
          
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row[0]) continue;
            
            const cellQuality = row.toString().trim();
            const cellQualityUpper = cellQuality.toUpperCase();
            const cellQualityLower = cellQuality.toLowerCase();
            
            if (cellQuality === qualityOriginal || 
                cellQualityUpper === qualityUpper || 
                cellQualityLower === qualityLower ||
                cellQuality.includes(qualityOriginal) ||
                cellQualityUpper.includes(qualityUpper) ||
                cellQualityLower.includes(qualityLower)) {
              
              const stockValue = row[4] ? row[3].toString().trim() : '0';
              console.log(`FOUND: ${searchQuality} in ${file.name}: ${stockValue}`);
              results[searchQuality][file.name] = stockValue;
              break;
            }
          }
        });

      } catch (sheetError) {
        console.error(`Error accessing stock sheet ${file.name}:`, sheetError.message);
      }
    }

    return results;

  } catch (error) {
    console.error('Error searching stock sheets:', error);
    throw error;
  }
}

async function handleMultipleOrderSelectionWithHiddenField(from, userInput, productId, phoneId) {
  try {
    const userState = userStates[from];
    if (!userState) {
      await sendWhatsAppMessage(from, '❌ Session expired. Start over: /→3', productId, phoneId);
      return;
    }
    
    if (/^\d+$/.test(userInput.trim())) {
      // Single store for all items
      const storeIndex = parseInt(userInput) - 1;
      const selectedStore = userState.permittedStores[storeIndex];
      
      if (selectedStore) {
        await createSingleStoreForm(from, selectedStore, userState.qualities, productId, phoneId);
      } else {
        await sendWhatsAppMessage(from, `❌ Invalid store number`, productId, phoneId);
      }
    } else {
      // Multiple store combinations
      const combinations = parseMultipleOrderInput(userInput, userState);
      
      if (combinations.length > 0) {
        await createMultipleStoreForms(from, combinations, productId, phoneId);
      } else {
        await sendWhatsAppMessage(from, '❌ Invalid format. Try: Quality-StoreNumber, Quality-StoreNumber', productId, phoneId);
      }
    }
    
  } catch (error) {
    console.error('Error handling multiple orders:', error);
  }
}

function parseMultipleOrderInput(input, userState) {
  try {
    const combinations = [];
    const parts = input.split(',');
    
    parts.forEach(part => {
      const trimmed = part.trim();
      const match = trimmed.match(/^(.+)-(\d+)$/);
      
      if (match) {
        const quality = match[1].trim();
        const storeIndex = parseInt(match[4]) - 1;
        const store = userState.permittedStores[storeIndex];
        
        if (store && userState.qualities.includes(quality)) {
          combinations.push({ quality, store });
        }
      }
    });
    
    return combinations;
  } catch (error) {
    return [];
  }
}

async function createSingleStoreForm(from, selectedStore, qualities, productId, phoneId) {
  try {
    const cleanPhone = from.replace(/^\+/, '');
    
    const formUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url` +
      `&entry.740712049=${encodeURIComponent(cleanPhone)}` +
      `&store=${encodeURIComponent(selectedStore)}`;
    
    let confirmationMessage = `✅ *${selectedStore}*\n\n`;
    confirmationMessage += `📋 ${formUrl}\n\n`;
    confirmationMessage += `_Type */* for main menu._`;
    
    await sendWhatsAppMessage(from, confirmationMessage, productId, phoneId);
    
    userStates[from] = { currentMenu: 'completed' };
    
  } catch (error) {
    console.error('Error creating single store form:', error);
  }
}

async function createMultipleStoreForms(from, combinations, productId, phoneId) {
  try {
    const cleanPhone = from.replace(/^\+/, '');
    let responseMessage = `✅ *MULTIPLE ORDER FORMS*\n\n`;
    
    // Group by store
    const storeGroups = {};
    combinations.forEach(combo => {
      if (!storeGroups[combo.store]) {
        storeGroups[combo.store] = [];
      }
      storeGroups[combo.store].push(combo.quality);
    });
    
    // Create form for each store
    Object.entries(storeGroups).forEach(([store, qualities]) => {
      const formUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url` +
        `&entry.740712049=${encodeURIComponent(cleanPhone)}` +
        `&store=${encodeURIComponent(store)}`;
      
      responseMessage += `🏪 *${store}*\n`;
      responseMessage += `📦 ${qualities.join(', ')}\n`;
      responseMessage += `📋 ${formUrl}\n\n`;
    });
    
    responseMessage += `_Fill each form for your different store orders._`;
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
    userStates[from] = { currentMenu: 'completed' };
    
  } catch (error) {
    console.error('Error creating multiple forms:', error);
  }
}

async function sendWhatsAppMessage(to, message, productId, phoneId) {
  try {
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
    console.log('Message sent successfully');
  } catch (error) {
    console.error('Error sending message:', error.response?.data || error.message);
  }
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🤖 WhatsApp Bot running on port ${PORT}`);
  console.log('✅ Bot ready with COMPLETE Order Query + Stock Query functionality!');
  console.log(`📦 Live Sheet ID: ${LIVE_SHEET_ID} (FMS Sheet)`);
  console.log(`📁 Completed Order Folder ID: ${COMPLETED_ORDER_FOLDER_ID}`);
  console.log(`📊 Stock Folder ID: ${STOCK_FOLDER_ID}`);
  console.log('🎯 Both Order Query and Stock Query now fully implemented!');
});
