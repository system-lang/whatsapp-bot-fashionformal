require('dotenv').config();
const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.json());

// Add file serving endpoint for PDFs
app.get('/download/:filename', (req, res) => {
  try {
    const filename = req.params.filename;
    const filepath = path.join(__dirname, 'temp', filename);
    
    if (fs.existsSync(filepath)) {
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.download(filepath, filename, (err) => {
        if (err) {
          console.error('Download error:', err);
        } else {
          setTimeout(() => {
            try {
              fs.unlinkSync(filepath);
            } catch (cleanupErr) {
              // Silent cleanup
            }
          }, 300000);
        }
      });
    } else {
      res.status(404).send(`
        <html>
          <body>
            <h2>File Not Found</h2>
            <p>The requested PDF file was not found or has expired.</p>
            <p>Please generate a new query.</p>
          </body>
        </html>
      `);
    }
  } catch (error) {
    console.error('File serving error:', error);
    res.status(500).send('Error serving file');
  }
});

// Store user states to track which menu they're in
let userStates = {};
let orderQueryTimestamps = {};
const STOCK_SESSION_TIMEOUT = 40 * 1000; // 40 seconds in milliseconds

const SHORTCUT_COMMANDS = [
  '/menu', '/stock', '/order', '/shirting', '/jacket', '/trouser',
  '/helpticket', '/delegation', '/', '/debuggreet', '/debugpermissions', '/debugrows'
];

const links = {
  helpTicket: 'https://tinyurl.com/HelpticketFF',
  delegation: 'https://tinyurl.com/DelegationFF',
  leave: 'YOUR_LEAVE_FORM_LINK_HERE'
};

const MAYTAPI_API_TOKEN = '07d75e68-b94f-485b-9e8c-19e707d176ae';

// Google Sheets configuration
const STOCK_FOLDER_ID = '1QV1cJ9jJZZW2PY24uUY2hefKeUqVHrrf';
const STORE_PERMISSION_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';
const GREETINGS_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';
const USER_ACCESS_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';
const STATIC_FORM_BASE_URL = 'https://docs.google.com/forms/d/e/1FAIpQLSfyAo7LwYtQDfNVxPRbHdk_ymGpDs-RyWTCgzd2PdRhj0T3Hw/viewform';

// Shirting Order Configuration
const LIVE_SHEET_ID = '1AxjCHsMxYUmEULaW1LxkW78g0Bv9fp4PkZteJO82uEA';
const LIVE_SHEET_NAME = 'FMS';
const COMPLETED_ORDER_FOLDER_ID = '1kgdPdnUK-FsnKZDE5yW6vtRf2H9d3YRE';

// Jacket Order Configuration
const JACKET_LIVE_SHEET_ID = '1XYXOv6C-aIuMVYDLSMflPZIQL7yJWq5BgnmDAnRMt58';
const JACKET_LIVE_SHEET_NAME = 'FMS';
const JACKET_COMPLETED_ORDER_FOLDER_ID = '1GmcGommmEBlP4iNPRA6NC4nbyFokCdY8';

const PRODUCTION_STAGES = [
  { name: 'CUT', column: 'T', nextStage: 'FUS' },
  { name: 'FUS', column: 'Z', nextStage: 'PAS' },
  { name: 'PAS', column: 'AF', nextStage: 'MAK' },
  { name: 'MAK', column: 'AL', nextStage: 'BH' },
  { name: 'BH', column: 'AX', nextStage: 'BS' },
  { name: 'BS', column: 'BD', nextStage: 'QC' },
  { name: 'QC', column: 'BJ', nextStage: 'ALT' },
  { name: 'ALT', column: 'BP', nextStage: 'IRO' },
  { name: 'IRO', column: 'BY', nextStage: 'Dispatch (Factory)' },
  { name: 'Dispatch (Factory)', column: 'CE', nextStage: 'Dispatch (HO)' },
  { name: 'Dispatch (HO)', column: 'CL', nextStage: 'COMPLETED', dispatchDateColumn: 'CL' }
];

const JACKET_PRODUCTION_STAGES = [
  { name: 'CUT', column: 'T', nextStage: 'FUS' },
  { name: 'FUS', column: 'Z', nextStage: 'Prep' },
  { name: 'Prep', column: 'AF', nextStage: 'MAK' },
  { name: 'MAK', column: 'AL', nextStage: 'QC1' },
  { name: 'QC1', column: 'AR', nextStage: 'BH' },
  { name: 'BH', column: 'AX', nextStage: 'Press' },
  { name: 'Press', column: 'BD', nextStage: 'QC2' },
  { name: 'QC2', column: 'BJ', nextStage: 'Dispatch (Factory)' },
  { name: 'Dispatch (Factory)', column: 'BP', nextStage: 'Dispatch (HO)' },
  { name: 'Dispatch (HO)', column: 'BW', nextStage: 'COMPLETED', dispatchDateColumn: 'BW' }
];

// Helper functions
function isStockSessionExpired(userState) {
  if (!userState || userState.currentMenu !== 'smart_stock_query' || !userState.lastActivity) return false;
  return (Date.now() - userState.lastActivity) > STOCK_SESSION_TIMEOUT;
}

function updateLastActivity(from) {
  if (userStates[from] && userStates[from].currentMenu === 'smart_stock_query') {
    userStates[from].lastActivity = Date.now();
    console.log(`Updated stock session activity for ${from} at ${new Date().toLocaleTimeString()}`);
  }
}

function isValidBotInteraction(message, userState) {
  const lowerMessage = message.toLowerCase().trim();
  
  if (lowerMessage.startsWith('/debugorder ') || lowerMessage.startsWith('/debugjacket ') || SHORTCUT_COMMANDS.includes(lowerMessage) || ['1', '2', '3', '4'].includes(message.trim())) {
    return true;
  }
  
  if (userState) {
    switch (userState.currentMenu) {
      case 'main': return ['1', '2', '3', '4'].includes(message.trim());
      case 'order_query': return ['1', '2', '3'].includes(message.trim());
      case 'order_number_input':
      case 'order_followup': return message.trim().length > 0;
      case 'smart_stock_query': 
        if (isStockSessionExpired(userState)) return false;
        return message.trim().length > 0;
      default: return false;
    }
  }
  return false;
}

async function getGoogleAuth() {
  try {
    const base64Key = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
    if (!base64Key) throw new Error('GOOGLE_SERVICE_ACCOUNT_KEY environment variable not set');
    
    const keyBuffer = Buffer.from(base64Key, 'base64');
    const keyData = JSON.parse(keyBuffer.toString());
    
    return new google.auth.GoogleAuth({
      credentials: keyData,
      scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly'],
    });
  } catch (error) {
    console.error('Error setting up Google auth:', error);
    throw error;
  }
}

function formatDateForDisplay(rawDate) {
  if (!rawDate || rawDate === '') return 'Date not available';
  
  const dateStr = rawDate.toString().trim();
  if (dateStr.includes('/') || dateStr.includes('-') || dateStr.includes(' ')) return dateStr;
  
  const dateNum = parseFloat(dateStr);
  if (!isNaN(dateNum) && dateNum > 1000) {
    try {
      const jsDate = new Date((dateNum - 25569) * 86400 * 1000);
      if (!isNaN(jsDate.getTime())) {
        return jsDate.toLocaleString('en-IN', { 
          timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit'
        });
      }
    } catch (error) {}
  }
  return dateStr;
}

function formatStockQuantity(stockValue) {
  if (!stockValue || stockValue === '') return stockValue;
  const numValue = parseFloat(stockValue.toString().trim());
  if (!isNaN(numValue) && numValue > 15) return '15+';
  return stockValue.toString();
}

function goBackOneStep(from) {
  if (!userStates[from]) return false;
  
  const currentMenu = userStates[from].currentMenu;
  if (currentMenu === 'order_query') {
    userStates[from] = { currentMenu: 'main', timestamp: Date.now() };
    return true;
  }
  if (currentMenu === 'order_number_input') {
    userStates[from] = { currentMenu: 'order_query', timestamp: Date.now() };
    return true;
  }
  if (currentMenu === 'smart_stock_query') {
    userStates[from] = { currentMenu: 'main', timestamp: Date.now() };
    return true;
  }
  return false;
}

function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

function isWithinOrderQueryWindow(from) {
  if (!orderQueryTimestamps[from]) return false;
  const now = Date.now();
  const lastQuery = orderQueryTimestamps[from];
  const twoMinutes = 2 * 60 * 1000;
  return (now - lastQuery) < twoMinutes;
}

// FIXED: More flexible order matching function for jacket orders
function isOrderMatch(orderNumber, searchTerm) {
  const upperOrderNumber = orderNumber.toUpperCase();
  const upperSearchTerm = searchTerm.toUpperCase();
  
  // 1. Direct substring match (most common case)
  if (upperOrderNumber.includes(upperSearchTerm)) {
    return true;
  }
  
  // 2. Split by separators and check each part
  const orderParts = upperOrderNumber.split(/[-_\s#]/);
  
  for (const part of orderParts) {
    // Exact match
    if (part === upperSearchTerm) {
      return true;
    }
    // Starts with match (minimum 3 characters to avoid false positives)
    if (upperSearchTerm.length >= 3 && part.startsWith(upperSearchTerm)) {
      return true;
    }
  }
  
  // 3. For numeric searches, be more flexible but still avoid false positives
  if (/^\d+$/.test(upperSearchTerm) && upperSearchTerm.length >= 4) {
    // For pure numbers, check if it appears as a distinct number sequence
    const numberRegex = new RegExp(`\\b${upperSearchTerm}\\d*\\b`);
    if (numberRegex.test(upperOrderNumber)) {
      return true;
    }
  }
  
  return false;
}

async function getUserPermissions(phoneNumber) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: USER_ACCESS_SHEET_ID,
      range: "'BOT Permission'!A:B",
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return [];

    const phoneVariations = [
      phoneNumber, phoneNumber.replace(/^\+91/, ''), phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''), phoneNumber.replace(/^0/, ''), phoneNumber.slice(-10)
    ];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 2) continue;

      const sheetPhone = (row[0] || '').toString().trim();
      const featuresString = (row[1] || '').toString().trim();

      for (const phoneVar of phoneVariations) {
        if (phoneVar === sheetPhone) {
          return featuresString.split(',').map(f => f.trim().toLowerCase());
        }
      }
    }
    return [];
  } catch (error) {
    console.error('Error getting user permissions:', error);
    return [];
  }
}

async function hasFeatureAccess(phoneNumber, feature) {
  const userPermissions = await getUserPermissions(phoneNumber);
  return userPermissions.includes(feature.toLowerCase());
}

async function getUserPermittedStores(phoneNumber) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: STORE_PERMISSION_SHEET_ID,
      range: 'store permission!A:B',
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return [];
    
    const permittedStores = [];
    const phoneVariations = [
      phoneNumber, phoneNumber.replace(/^\+91/, ''), phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''), phoneNumber.replace(/^0/, ''), phoneNumber.slice(-10)
    ];
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 2) continue;
      
      const sheetContact = (row[0] || '').toString().trim();
      const sheetStore = (row[1] || '').toString().trim();
      
      for (const phoneVar of phoneVariations) {
        if (phoneVar === sheetContact) {
          if (sheetStore && sheetStore !== '') {
            permittedStores.push(sheetStore);
          }
          break;
        }
      }
    }
    return permittedStores;
  } catch (error) {
    console.error('Error getting permitted stores:', error);
    return [];
  }
}

async function getUserGreeting(phoneNumber) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const attempts = ['Greetings!A:D', "'Greetings'!A:D", 'Sheet2!A:D', 'A:D'];
    
    let rows = null;
    for (const range of attempts) {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: GREETINGS_SHEET_ID,
          range: range,
          valueRenderOption: 'UNFORMATTED_VALUE'
        });
        
        if (response.data.values && response.data.values.length > 0) {
          rows = response.data.values;
          break;
        }
      } catch (rangeError) { continue; }
    }
    
    if (!rows || rows.length <= 1) return null;

    const phoneVariations = [
      phoneNumber, phoneNumber.replace(/^\+91/, ''), phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''), phoneNumber.replace(/^0/, ''), phoneNumber.slice(-10)
    ];
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 4) continue;
      
      const sheetContact = (row[0] || '').toString().trim();
      const name = (row[1] || '').toString().trim();
      const salutation = (row[2] || '').toString().trim();
      const greetings = (row[3] || '').toString().trim();
      
      for (const phoneVar of phoneVariations) {
        if (phoneVar === sheetContact) {
          return { name, salutation, greetings };
        }
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting greeting:', error);
    return null;
  }
}

function formatGreetingMessage(greeting, mainMessage) {
  if (!greeting || !greeting.name || !greeting.salutation || !greeting.greetings) return mainMessage;
  return `${greeting.salutation} ${greeting.name}\n\n${greeting.greetings}\n\n${mainMessage}`;
}

async function generatePersonalizedMenu(phoneNumber) {
  const userPermissions = await getUserPermissions(phoneNumber);
  
  if (userPermissions.length === 0) {
    return `*ACCESS DENIED*\n\nYou do not have permission to use this bot.\nPlease contact administrator for access.`;
  }

  let menuItems = [];
  let shortcuts = [];

  const hasAnyTicketAccess = userPermissions.some(perm => 
    ['help_ticket', 'delegation', 'leave_form'].includes(perm)
  );

  if (hasAnyTicketAccess) {
    menuItems.push('1. Ticket');
    if (userPermissions.includes('help_ticket')) shortcuts.push('/helpticket - Direct Help Ticket');
    if (userPermissions.includes('delegation')) shortcuts.push('/delegation - Direct Delegation');
  }
  
  if (userPermissions.includes('order')) {
    menuItems.push('2. Order Query');
    shortcuts.push('/order - Direct Order Menu', '/shirting - Shirting Orders', '/jacket - Jacket Orders', '/trouser - Trouser Orders');
  }
  
  if (userPermissions.includes('stock')) {
    menuItems.push('3. Stock Query');
    shortcuts.push('/stock - Direct Stock Query');
  }
  
  if (userPermissions.includes('document')) menuItems.push('4. Document');

  if (menuItems.length === 0) {
    return `*ACCESS DENIED*\n\nNo valid permissions found.\nPlease contact administrator for access.`;
  }

  let menu = `*MAIN MENU*\n\nPlease select an option:\n\n${menuItems.join('\n')}`;
  if (shortcuts.length > 0) menu += `\n\n*SHORTCUTS:*\n${shortcuts.join('\n')}`;
  menu += `\n\nType the number, use shortcuts, or / to go back`;
  return menu;
}

async function generateTicketMenu(phoneNumber) {
  const userPermissions = await getUserPermissions(phoneNumber);
  let ticketOptions = [];
  
  if (userPermissions.includes('help_ticket')) ticketOptions.push(`*HELP TICKET*\n${links.helpTicket}`);
  if (userPermissions.includes('delegation')) ticketOptions.push(`*DELEGATION*\n${links.delegation}`);
  if (userPermissions.includes('leave_form')) ticketOptions.push(`*LEAVE FORM*\n${links.leave}`);

  if (ticketOptions.length === 0) {
    return `*ACCESS DENIED*\n\nYou don't have permission to access any ticket options.\nContact administrator for access.\n\nType /menu to return to main menu`;
  }

  return `*TICKET OPTIONS*\n\nClick the links below to access forms directly:\n\n${ticketOptions.join('\n\n')}\n\nType /menu to return to main menu or / to go back`;
}

function checkProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let hasAnyStage = false;

    for (let i = 0; i < PRODUCTION_STAGES.length; i++) {
      const stage = PRODUCTION_STAGES[i];
      const columnIndex = columnToIndex(stage.column);
      
      let cellValue = '';
      if (row.length > columnIndex && row[columnIndex] !== undefined && row[columnIndex] !== null) {
        cellValue = row[columnIndex].toString().trim();
      }
      
      if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
        lastCompletedStage = stage;
        hasAnyStage = true;
      }
    }

    if (!hasAnyStage) return { message: 'Order is currently under process' };

    if (lastCompletedStage && lastCompletedStage.name === 'Dispatch (HO)') {
      const dispatchDateIndex = columnToIndex(lastCompletedStage.dispatchDateColumn);
      let rawDispatchDate = '';
      if (row.length > dispatchDateIndex && row[dispatchDateIndex] !== undefined && row[dispatchDateIndex] !== null) {
        rawDispatchDate = row[dispatchDateIndex];
      }
      const formattedDate = formatDateForDisplay(rawDispatchDate);
      return { message: `Order has been dispatched from HO on ${formattedDate}` };
    }

    if (lastCompletedStage) {
      return { message: `Order is currently completed ${lastCompletedStage.name} stage and processed to ${lastCompletedStage.nextStage} stage` };
    }

    return { message: 'Error determining order status' };
  } catch (error) {
    console.error('Error checking production stages:', error);
    return { message: 'Error checking order status' };
  }
}

function checkJacketProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let hasAnyStage = false;

    for (let i = 0; i < JACKET_PRODUCTION_STAGES.length; i++) {
      const stage = JACKET_PRODUCTION_STAGES[i];
      const columnIndex = columnToIndex(stage.column);
      
      let cellValue = '';
      if (row.length > columnIndex && row[columnIndex] !== undefined && row[columnIndex] !== null) {
        cellValue = row[columnIndex].toString().trim();
      }
      
      if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
        lastCompletedStage = stage;
        hasAnyStage = true;
      }
    }

    if (!hasAnyStage) return { message: 'Order is currently under process' };

    if (lastCompletedStage && lastCompletedStage.name === 'Dispatch (HO)') {
      const dispatchDateIndex = columnToIndex(lastCompletedStage.dispatchDateColumn);
      let rawDispatchDate = '';
      if (row.length > dispatchDateIndex && row[dispatchDateIndex] !== undefined && row[dispatchDateIndex] !== null) {
        rawDispatchDate = row[dispatchDateIndex];
      }
      const formattedDate = formatDateForDisplay(rawDispatchDate);
      return { message: `Order has been dispatched from HO on ${formattedDate}` };
    }

    if (lastCompletedStage) {
      return { message: `Order is currently completed ${lastCompletedStage.name} stage and processed to ${lastCompletedStage.nextStage} stage` };
    }

    return { message: 'Error determining order status' };
  } catch (error) {
    console.error('Error checking jacket production stages:', error);
    return { message: 'Error checking order status' };
  }
}

async function searchStockWithPartialMatch(searchTerms) {
  const results = {};
  
  searchTerms.forEach(term => {
    results[term] = [];
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

    for (const file of folderFiles.data.files) {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: file.id,
          range: 'A:E',
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) continue;
        
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length < 5) continue;
          
          const qualityCode = (row[0] || '').toString().trim();
          const stockQuantity = (row[4] || '').toString().trim();
          
          if (qualityCode && stockQuantity && qualityCode !== '' && stockQuantity !== '') {
            searchTerms.forEach(searchTerm => {
              const cleanSearchTerm = searchTerm.trim();
              
              if (cleanSearchTerm.length >= 5 && qualityCode.toUpperCase().includes(cleanSearchTerm.toUpperCase())) {
                results[searchTerm].push({
                  qualityCode: qualityCode,
                  stock: stockQuantity,
                  store: file.name,
                  searchTerm: cleanSearchTerm
                });
              }
            });
          }
        }

      } catch (sheetError) {
        console.error(`Error processing ${file.name}:`, sheetError.message);
      }
    }

    // Remove duplicates and keep maximum stock quantity for each quality code
    searchTerms.forEach(term => {
      if (results[term] && results[term].length > 0) {
        const qualityCodeMap = {};
        
        results[term].forEach(item => {
          const code = item.qualityCode;
          const stockNum = parseFloat(item.stock) || 0;
          
          if (!qualityCodeMap[code] || stockNum > parseFloat(qualityCodeMap[code].stock)) {
            qualityCodeMap[code] = item;
          }
        });
        
        results[term] = Object.values(qualityCodeMap);
      }
    });

    return results;

  } catch (error) {
    console.error('Stock search error:', error);
    throw error;
  }
}

async function processSmartStockQuery(from, searchTerms, productId, phoneId) {
  try {
    updateLastActivity(from);
    
    const validTerms = searchTerms.filter(term => {
      const cleanTerm = term.trim();
      return cleanTerm.length >= 5;
    });
    
    if (validTerms.length === 0) {
      await sendWhatsAppMessage(from, `*Invalid Search*\n\nPlease provide at least 5 characters for searching.\n\nExamples:\n- 11010 (finds 11010088471-001)\n- ABC12 (finds ABC123456789)\n- 88471 (finds 11010088471-001)\n\nYou can search again within 40 seconds or type /menu for main menu`, productId, phoneId);
      
      userStates[from] = { 
        currentMenu: 'smart_stock_query',
        lastActivity: Date.now()
      };
      return;
    }
    
    await sendWhatsAppMessage(from, `*Smart Stock Search*\n\nSearching for: ${validTerms.join(', ')}\n\nPlease wait while I search all stock sheets...`, productId, phoneId);

    const searchResults = await searchStockWithPartialMatch(validTerms);
    const permittedStores = await getUserPermittedStores(from);
    
    let totalResults = 0;
    validTerms.forEach(term => {
      totalResults += (searchResults[term] || []).length;
    });
    
    if (totalResults === 0) {
      const noResultsMessage = `*No Results Found*\n\nNo stock items found containing:\n${validTerms.map(term => `- ${term}`).join('\n')}\n\nTry:\n- Different search combinations\n- Shorter terms (5+ characters)\n- Both letters and numbers work\n\nYou can search again within 40 seconds or type /menu for main menu`;
      
      await sendWhatsAppMessage(from, noResultsMessage, productId, phoneId);
      
      userStates[from] = { 
        currentMenu: 'smart_stock_query',
        lastActivity: Date.now()
      };
      return;
    }
    
    if (totalResults <= 15) {
      let responseMessage = `*Smart Search Results*\n\n`;
      
      const storeGroups = {};
      validTerms.forEach(term => {
        const termResults = searchResults[term] || [];
        termResults.forEach(result => {
          if (!storeGroups[result.store]) {
            storeGroups[result.store] = [];
          }
          storeGroups[result.store].push(result);
        });
      });
      
      Object.entries(storeGroups).forEach(([store, items]) => {
        responseMessage += `*${store}*\n`;
        items.forEach(item => {
          const formattedStock = formatStockQuantity(item.stock);
          responseMessage += `${item.qualityCode}: ${formattedStock}\n`;
        });
        responseMessage += `\n`;
      });
      
      responseMessage += `*Place Orders*\n\n`;

      if (permittedStores.length === 0) {
        responseMessage += `No store permissions found\nContact admin for access\n\n`;
      } else if (permittedStores.length === 1) {
        const cleanPhone = from.replace(/^\+/, '');
        const formUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url&entry.740712049=${encodeURIComponent(cleanPhone)}&store=${encodeURIComponent(permittedStores[0])}`;
        responseMessage += `*Your Store:* ${permittedStores}\n${formUrl}\n\n`;
      } else {
        responseMessage += `*Your Stores:*\n`;
        permittedStores.forEach((store, index) => {
          responseMessage += `${index + 1}. ${store}\n`;
        });
        responseMessage += `\nReply with store number to get order form\n\n`;
      }
      
      responseMessage += `Search more items within 40 seconds or type /menu for main menu`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      
      userStates[from] = { 
        currentMenu: 'smart_stock_query',
        lastActivity: Date.now()
      };
      
    } else {
      try {
        const pdfResult = await generateStockPDF(searchResults, validTerms, from, permittedStores);
        
        const summaryMessage = `*Large Results Found*\n\nSearch: ${validTerms.join(', ')}\nTotal Results: ${totalResults} items\nPDF Generated: ${pdfResult.filename}\n\nResults are too long for WhatsApp\nPDF does NOT include order forms—see WhatsApp message for order form link\n\nSearch more items within 40 seconds or type /menu for main menu`;
        
        await sendWhatsAppMessage(from, summaryMessage, productId, phoneId);
        await sendWhatsAppFile(from, pdfResult.filepath, pdfResult.filename, productId, phoneId, permittedStores);
        
        userStates[from] = { 
          currentMenu: 'smart_stock_query',
          lastActivity: Date.now()
        };
        
      } catch (pdfError) {
        console.error('PDF generation failed:', pdfError);
        await sendWhatsAppMessage(from, `*Error Generating PDF*\n\nFound ${totalResults} results but could not generate PDF.\nPlease contact support.\n\nType /menu for main menu`, productId, phoneId);
        
        delete userStates[from];
      }
    }
    
  } catch (error) {
    console.error('Error in smart stock query:', error);
    await sendWhatsAppMessage(from, `*Search Error*\n\nUnable to complete search.\nPlease try again later.\n\nType /menu for main menu`, productId, phoneId);
    
    delete userStates[from];
  }
}

async function searchOrdersByPartialMatch(searchTerm) {
  const matchingOrders = [];
  
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    const cleanSearchTerm = searchTerm.trim().toUpperCase();
    
    // Search in live sheet first
    const liveResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: LIVE_SHEET_ID,
      range: `${LIVE_SHEET_NAME}!A:CL`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const liveRows = liveResponse.data.values;
    if (liveRows && liveRows.length > 1) {
      for (let i = 1; i < liveRows.length; i++) {
        const row = liveRows[i];
        if (!row || row.length === 0) continue;
        
        let orderNumber = '';
        if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
          orderNumber = row[3].toString().trim();
        }
        
        if (orderNumber && isOrderMatch(orderNumber, cleanSearchTerm)) {
          const stageStatus = checkProductionStages(row);
          matchingOrders.push({
            orderNumber: orderNumber,
            message: stageStatus.message,
            location: 'Live Sheet (Shirting FMS)'
          });
        }
      }
    }

    // Search in completed orders
    const folderFiles = await drive.files.list({
      q: `'${COMPLETED_ORDER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    for (const file of folderFiles.data.files) {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: file.id,
          range: 'A:CL',
          valueRenderOption: 'UNFORMATTED_VALUE'
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) continue;

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          
          let orderNumber = '';
          if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
            orderNumber = row[3].toString().trim();
          }
          
          if (orderNumber && isOrderMatch(orderNumber, cleanSearchTerm)) {
            let rawDispatchDate = '';
            if (row.length > 90 && row[90] !== undefined && row[90] !== null) {
              rawDispatchDate = row[90];
            }
            
            const formattedDate = formatDateForDisplay(rawDispatchDate);
            
            matchingOrders.push({
              orderNumber: orderNumber,
              message: `Order got dispatched on ${formattedDate}`,
              location: 'Completed Orders (Shirting)'
            });
          }
        }
      } catch (error) {
        continue;
      }
    }

    return matchingOrders;
  } catch (error) {
    console.error('Error in searchOrdersByPartialMatch:', error);
    return [];
  }
}

// FIXED: Enhanced jacket order search - searches ALL rows properly
async function searchJacketOrdersByPartialMatch(searchTerm) {
  const matchingOrders = [];
  
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    const cleanSearchTerm = searchTerm.trim().toUpperCase();
    
    console.log(`🔍 Searching for jacket orders matching: "${cleanSearchTerm}"`);

    // Search in jacket live sheet - GET ALL ROWS not just first 10
    try {
      const liveResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: JACKET_LIVE_SHEET_ID,
        range: `${JACKET_LIVE_SHEET_NAME}!A:BW`, // This gets ALL rows
        valueRenderOption: 'FORMATTED_VALUE'
      });

      const liveRows = liveResponse.data.values;
      if (liveRows && liveRows.length > 0) {
        console.log(`Searching ${liveRows.length} rows in jacket live sheet...`);
        
        // Start from row 2 and search ALL rows
        for (let i = 2; i < liveRows.length; i++) {
          const row = liveRows[i];
          if (!row || row.length === 0) continue;
          
          // Check column D (index 3) for order numbers
          let orderNumber = '';
          
          if (row.length > 3 && row[3] !== undefined && row[3] !== null && row[3] !== '') {
            orderNumber = row[3].toString().trim();
          }
          
          console.log(`Checking row ${i}: "${orderNumber}" vs "${cleanSearchTerm}"`);
          
          if (orderNumber && orderNumber !== '' && isOrderMatch(orderNumber, cleanSearchTerm)) {
            console.log(`✅ Found jacket match in live sheet: ${orderNumber}`);
            const stageStatus = checkJacketProductionStages(row);
            matchingOrders.push({
              orderNumber: orderNumber,
              message: stageStatus.message,
              location: 'Live Sheet (Jacket FMS)'
            });
          }
        }
      }
    } catch (liveSheetError) {
      console.error('Error searching jacket live sheet:', liveSheetError.message);
    }

    // Search in jacket completed orders
    try {
      const folderFiles = await drive.files.list({
        q: `'${JACKET_COMPLETED_ORDER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
        fields: 'files(id, name)'
      });

      console.log(`Searching ${folderFiles.data.files.length} jacket completed order files...`);

      for (const file of folderFiles.data.files) {
        try {
          const response = await sheets.spreadsheets.values.get({
            spreadsheetId: file.id,
            range: 'A:BW', // Get ALL rows
            valueRenderOption: 'FORMATTED_VALUE'
          });

          const rows = response.data.values;
          if (!rows || rows.length === 0) continue;

          // Start from row 2 and search ALL rows
          for (let i = 2; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            
            let orderNumber = '';
            
            // Check column D (index 3) for order numbers
            if (row.length > 3 && row[3] !== undefined && row[3] !== null && row[3] !== '') {
              orderNumber = row[3].toString().trim();
            }
            
            if (orderNumber && orderNumber !== '' && isOrderMatch(orderNumber, cleanSearchTerm)) {
              console.log(`✅ Found jacket match in completed orders: ${orderNumber}`);
              let rawDispatchDate = '';
              if (row.length > 74 && row[74] !== undefined && row[74] !== null) {
                rawDispatchDate = row[74];
              }
              
              const formattedDate = formatDateForDisplay(rawDispatchDate);
              
              matchingOrders.push({
                orderNumber: orderNumber,
                message: `Order got dispatched on ${formattedDate}`,
                location: 'Completed Orders (Jacket)'
              });
            }
          }
        } catch (error) {
          console.error(`Error searching jacket file ${file.name}:`, error.message);
          continue;
        }
      }
    } catch (folderError) {
      console.error('Error accessing jacket completed order folder:', folderError.message);
    }

    console.log(`📊 Total jacket matches found: ${matchingOrders.length}`);
    return matchingOrders;

  } catch (error) {
    console.error('Error in searchJacketOrdersByPartialMatch:', error);
    return [];
  }
}


async function searchInLiveSheet(sheets, orderNumber) {
  try {
    console.log(`=== SEARCHING FOR SHIRTING ORDER: ${orderNumber} ===`);
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: LIVE_SHEET_ID,
      range: `${LIVE_SHEET_NAME}!A:CL`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log('ERROR: No rows returned from shirting live sheet');
      return { found: false };
    }

    console.log(`Shirting live sheet has ${rows.length} total rows`);

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      
      let sheetOrderNumber = '';
      if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
        sheetOrderNumber = row[3].toString().trim();
      }
      
      if (sheetOrderNumber) {
        const searchOrder = orderNumber.trim().toUpperCase();
        const sheetOrder = sheetOrderNumber.toUpperCase();
        
        if (sheetOrder === searchOrder) {
          console.log(`✅ SHIRTING EXACT MATCH FOUND at row ${i}: ${sheetOrderNumber}`);
          const stageStatus = checkProductionStages(row);
          return {
            found: true,
            message: stageStatus.message,
            location: 'Live Sheet (Shirting FMS)'
          };
        }
      }
    }

    console.log(`❌ Shirting Order ${orderNumber} not found in any of the ${rows.length} rows`);
    return { found: false };

  } catch (error) {
    console.error('Error searching shirting live sheet:', error.message);
    return { found: false };
  }
}

// FIXED: Enhanced exact jacket search
async function searchInJacketLiveSheet(sheets, orderNumber) {
  try {
    console.log(`=== SEARCHING JACKET ORDER: ${orderNumber} ===`);
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: JACKET_LIVE_SHEET_ID,
      range: `${JACKET_LIVE_SHEET_NAME}!A:BW`,
      valueRenderOption: 'FORMATTED_VALUE'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log('ERROR: No rows returned from jacket live sheet');
      return { found: false };
    }

    console.log(`Jacket live sheet has ${rows.length} total rows`);

    // Try different starting points as sheets may have different header configurations
    const startRow = rows.length > 2 ? 1 : 0;

    for (let i = startRow; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      
      let sheetOrderNumber = '';
      
      // Try multiple columns for order numbers
      if (row.length > 4 && row[4] !== undefined && row[4] !== null && row[4] !== '') {
        sheetOrderNumber = row[4].toString().trim(); // Column E (index 4)
      } else if (row.length > 3 && row[3] !== undefined && row[3] !== null && row[3] !== '') {
        sheetOrderNumber = row[3].toString().trim(); // Column D (index 3)
      }
      
      if (sheetOrderNumber && sheetOrderNumber !== '') {
        const searchOrder = orderNumber.trim().toUpperCase();
        const sheetOrder = sheetOrderNumber.toUpperCase();
        
        if (sheetOrder === searchOrder) {
          console.log(`✅ JACKET EXACT MATCH FOUND at row ${i}: ${sheetOrderNumber}`);
          const stageStatus = checkJacketProductionStages(row);
          return {
            found: true,
            message: stageStatus.message,
            location: 'Live Sheet (Jacket FMS)'
          };
        }
      }
    }

    console.log(`❌ Jacket Order ${orderNumber} not found in any of the ${rows.length} rows`);
    return { found: false };

  } catch (error) {
    console.error('Error searching jacket live sheet:', error.message);
    return { found: false };
  }
}

async function searchInCompletedSheetSimplified(sheets, sheetId, orderNumber) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'A:CL',
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return { found: false };

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      
      let sheetOrderNumber = '';
      if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
        sheetOrderNumber = row[3].toString().trim();
      }
      
      if (!sheetOrderNumber) continue;

      if (sheetOrderNumber.toUpperCase() === orderNumber.trim().toUpperCase()) {
        let rawDispatchDate = '';
        if (row.length > 90 && row[90] !== undefined && row[90] !== null) {
          rawDispatchDate = row[90];
        }
        
        const formattedDate = formatDateForDisplay(rawDispatchDate);
        
        return {
          found: true,
          message: `Order got dispatched on ${formattedDate}`,
          location: 'Completed Orders (Shirting)'
        };
      }
    }

    return { found: false };

  } catch (error) {
    console.error('Error searching shirting completed sheet:', error);
    return { found: false };
  }
}

async function searchInJacketCompletedSheet(sheets, sheetId, orderNumber) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'A:BW',
      valueRenderOption: 'FORMATTED_VALUE'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return { found: false };

    const startRow = rows.length > 2 ? 1 : 0;

    for (let i = startRow; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      
      let sheetOrderNumber = '';
      
      // Try column E first (index 4)
      if (row.length > 4 && row[4] !== undefined && row[4] !== null && row[4] !== '') {
        sheetOrderNumber = row[4].toString().trim();
      }
      // Fallback to column D (index 3)
      else if (row.length > 3 && row[3] !== undefined && row[3] !== null && row[3] !== '') {
        sheetOrderNumber = row[3].toString().trim();
      }
      
      if (!sheetOrderNumber) continue;

      if (sheetOrderNumber.toUpperCase() === orderNumber.trim().toUpperCase()) {
        // Get dispatch date from BW column (index 74)
        let rawDispatchDate = '';
        if (row.length > 74 && row[74] !== undefined && row[74] !== null) {
          rawDispatchDate = row[74];
        }
        
        const formattedDate = formatDateForDisplay(rawDispatchDate);
        
        return {
          found: true,
          message: `Order got dispatched on ${formattedDate}`,
          location: 'Completed Orders (Jacket)'
        };
      }
    }

    return { found: false };

  } catch (error) {
    console.error('Error searching jacket completed sheet:', error);
    return { found: false };
  }
}

async function searchOrderStatus(orderNumber, category) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    const liveSheetResult = await searchInLiveSheet(sheets, orderNumber);
    if (liveSheetResult.found) return liveSheetResult;

    const folderFiles = await drive.files.list({
      q: `'${COMPLETED_ORDER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    for (const file of folderFiles.data.files) {
      try {
        const completedResult = await searchInCompletedSheetSimplified(sheets, file.id, orderNumber);
        if (completedResult.found) return completedResult;
      } catch (error) {
        continue;
      }
    }

    return { 
      found: false, 
      message: 'Order not found in system. Please contact responsible person.\n\nThank you.' 
    };

  } catch (error) {
    console.error('Error in searchOrderStatus:', error);
    return { 
      found: false, 
      message: 'Error occurred while searching order. Please contact responsible person.\n\nThank you.' 
    };
  }
}

async function searchJacketOrderStatus(orderNumber, category) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    const liveSheetResult = await searchInJacketLiveSheet(sheets, orderNumber);
    if (liveSheetResult.found) return liveSheetResult;

    const folderFiles = await drive.files.list({
      q: `'${JACKET_COMPLETED_ORDER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    for (const file of folderFiles.data.files) {
      try {
        const completedResult = await searchInJacketCompletedSheet(sheets, file.id, orderNumber);
        if (completedResult.found) return completedResult;
      } catch (error) {
        continue;
      }
    }

    return { 
      found: false, 
      message: 'Order not found in system. Please contact responsible person.\n\nThank you.' 
    };

  } catch (error) {
    console.error('Error in searchJacketOrderStatus:', error);
    return { 
      found: false, 
      message: 'Error occurred while searching order. Please contact responsible person.\n\nThank you.' 
    };
  }
}

async function generateStockPDF(searchResults, searchTerms, phoneNumber, permittedStores) {
  try {
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    const filename = `stock_results_${phoneNumber.slice(-4)}_${timestamp}.pdf`;
    const filepath = path.join(__dirname, 'temp', filename);
    
    const tempDir = path.join(__dirname, 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    const doc = new PDFDocument({
      margin: 50,
      size: 'A4'
    });

    const stream = fs.createWriteStream(filepath);
    doc.pipe(stream);

    doc.fontSize(18)
       .font('Helvetica-Bold')
       .text('STOCK QUERY RESULTS', { align: 'center' });
    
    doc.moveDown(0.5);
    
    doc.fontSize(10)
       .font('Helvetica')
       .text(`Generated: ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}`)
       .text(`Search Terms: ${searchTerms.join(', ')}`)
       .text(`Phone: ${phoneNumber}`)
       .moveDown();

    doc.moveTo(50, doc.y)
       .lineTo(550, doc.y)
       .stroke();
    doc.moveDown(0.5);

    let totalResults = 0;
    
    const allStoreGroups = {};
    
    searchTerms.forEach(term => {
      const termResults = searchResults[term] || [];
      totalResults += termResults.length;
      
      termResults.forEach(result => {
        if (!allStoreGroups[result.store]) {
          allStoreGroups[result.store] = [];
        }
        allStoreGroups[result.store].push({
          qualityCode: result.qualityCode,
          stock: formatStockQuantity(result.stock)
        });
      });
    });

    Object.entries(allStoreGroups).forEach(([storeName, items]) => {
      doc.fontSize(16)
         .font('Helvetica-Bold')
         .text(storeName);
      
      doc.moveDown(0.3);
      
      doc.fontSize(11)
         .font('Helvetica-Bold')
         .text('Quality Code', 50, doc.y, { width: 300, continued: true })
         .text('Stock Quantity', 350, doc.y);
      
      doc.moveDown(0.2);
      
      doc.moveTo(50, doc.y)
         .lineTo(500, doc.y)
         .stroke();
      doc.moveDown(0.3);
      
      doc.fontSize(10)
         .font('Helvetica');
      
      items.forEach(item => {
        doc.text(item.qualityCode, 50, doc.y, { width: 300, continued: true })
           .text(item.stock, 350, doc.y);
        doc.moveDown(0.15);
      });
      
      doc.moveDown(0.5);
    });

    if (totalResults === 0) {
      doc.fontSize(12)
         .font('Helvetica')
         .text('No matching results found.', { align: 'center' });
    } else {
      doc.fontSize(8)
         .font('Helvetica')
         .text(`Total Results: ${totalResults}`, { align: 'right' });
    }

    doc.end();

    await new Promise((resolve, reject) => {
      stream.on('finish', resolve);
      stream.on('error', reject);
    });

    return { filepath, filename, totalResults };

  } catch (error) {
    console.error('Error generating PDF:', error);
    throw error;
  }
}

async function generateOrderPDF(orderResults, searchTerm, phoneNumber, category) {
  try {
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    const filename = `order_results_${phoneNumber.slice(-4)}_${timestamp}.pdf`;
    const filepath = path.join(__dirname, 'temp', filename);
    
    const tempDir = path.join(__dirname, 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    const doc = new PDFDocument({
      margin: 50,
      size: 'A4'
    });

    const stream = fs.createWriteStream(filepath);
    doc.pipe(stream);

    doc.fontSize(18)
       .font('Helvetica-Bold')
       .text('ORDER QUERY RESULTS', { align: 'center' });
    
    doc.moveDown(0.5);
    
    doc.fontSize(10)
       .font('Helvetica')
       .text(`Generated: ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}`)
       .text(`Search Term: ${searchTerm}`)
       .text(`Category: ${category}`)
       .text(`Phone: ${phoneNumber}`)
       .moveDown();

    doc.moveTo(50, doc.y)
       .lineTo(550, doc.y)
       .stroke();
    doc.moveDown(0.5);

    doc.fontSize(11)
       .font('Helvetica-Bold')
       .text('Order Number', 50, doc.y, { width: 200, continued: true })
       .text('Status', 250, doc.y);
    
    doc.moveDown(0.2);
    
    doc.moveTo(50, doc.y)
       .lineTo(500, doc.y)
       .stroke();
    doc.moveDown(0.3);
    
    doc.fontSize(10)
       .font('Helvetica');
    
    orderResults.forEach(result => {
      doc.text(result.orderNumber, 50, doc.y, { width: 200, continued: true })
         .text(result.message.replace(/\*/g, ''), 250, doc.y);
      doc.moveDown(0.3);
    });

    doc.fontSize(8)
       .font('Helvetica')
       .text(`Total Results: ${orderResults.length}`, { align: 'right' });

    doc.end();

    await new Promise((resolve, reject) => {
      stream.on('finish', resolve);
      stream.on('error', reject);
    });

    return { filepath, filename, totalResults: orderResults.length };

  } catch (error) {
    console.error('Error generating order PDF:', error);
    throw error;
  }
}

async function sendWhatsAppFile(to, filepath, filename, productId, phoneId, permittedStores) {
  try {
    const fileStats = fs.statSync(filepath);
    const fileSizeKB = Math.round(fileStats.size / 1024);

    const baseUrl = 'https://whatsapp-bot-fashionformal-production.up.railway.app';
    const downloadUrl = `${baseUrl}/download/${filename}`;

    let orderFormMessage = '';
    if (permittedStores && permittedStores.length > 0) {
      if (permittedStores.length === 1) {
        const cleanPhone = to.replace(/^\+/, '');
        const formUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url&entry.740712049=${encodeURIComponent(cleanPhone)}&store=${encodeURIComponent(permittedStores[0])}`;
        orderFormMessage = `\n\n*Order Form Link:*\n${formUrl}\n`;
      } else {
        orderFormMessage = `\n\n*Order Form Link:*\nReply with your store number to get the order form link.`;
      }
    }

    const message = `*Results Generated*\n\nFile: ${filename}\nSize: ${fileSizeKB} KB\n\nDownload your PDF:\n${downloadUrl}${orderFormMessage}\n\nClick the link above to download\nWorks on mobile and desktop\nLink expires in 5 minutes\n\nType /menu for main menu`;

    await sendWhatsAppMessage(to, message, productId, phoneId);

  } catch (error) {
    console.error('Error creating download link:', error);
  }
}

async function processOrderQuery(from, category, orderNumbers, productId, phoneId, isFollowUp = false) {
  try {
    if (!isFollowUp) {
      await sendWhatsAppMessage(from, `*Checking ${category} orders...*\n\nPlease wait while I search for your order status.`, productId, phoneId);
    }

    let orderResults = [];
    let exactMatches = [];
    let partialMatches = [];

    for (const orderInput of orderNumbers) {
      const cleanInput = orderInput.trim();
      
      console.log(`Processing ${category} input: "${cleanInput}"`);
      
      if (cleanInput.length >= 3 && cleanInput.length <= 10) {
        console.log(`Performing partial search for ${category}: ${cleanInput}`);
        
        let matchingOrders = [];
        if (category === 'Jacket') {
          matchingOrders = await searchJacketOrdersByPartialMatch(cleanInput);
        } else {
          matchingOrders = await searchOrdersByPartialMatch(cleanInput);
        }
        
        partialMatches = partialMatches.concat(matchingOrders);
        
      } else if (cleanInput.length > 10) {
        console.log(`Performing exact search for ${category}: ${cleanInput}`);
        
        let orderStatus;
        if (category === 'Jacket') {
          orderStatus = await searchJacketOrderStatus(cleanInput, category);
        } else {
          orderStatus = await searchOrderStatus(cleanInput, category);
        }
        
        exactMatches.push({
          orderNumber: cleanInput,
          message: orderStatus.message,
          location: orderStatus.location || 'Unknown'
        });
      }
    }

    // Combine all results and remove duplicates
    const allResults = [...exactMatches, ...partialMatches];
    const uniqueResults = [];
    const seenOrders = new Set();
    
    for (const result of allResults) {
      if (!seenOrders.has(result.orderNumber)) {
        seenOrders.add(result.orderNumber);
        uniqueResults.push(result);
      }
    }
    
    orderResults = uniqueResults;

    console.log(`Final ${category} results count: ${orderResults.length}`);

    if (orderResults.length === 0) {
      let responseMessage = isFollowUp ? `*ADDITIONAL ${category.toUpperCase()} ORDER STATUS*\n\n` : `*${category.toUpperCase()} ORDER STATUS*\n\n`;
      responseMessage += `No orders found for the given search terms.\n\nType /menu for main menu or / to go back`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      return;
    }

    if (orderResults.length <= 3) {
      let responseMessage = isFollowUp ? `*ADDITIONAL ${category.toUpperCase()} ORDER STATUS*\n\n` : `*${category.toUpperCase()} ORDER STATUS*\n\n`;
      
      orderResults.forEach(result => {
        responseMessage += `*Order: ${result.orderNumber}*\n${result.message}\n\n`;
      });
      
      orderQueryTimestamps[from] = Date.now();
      
      responseMessage += `You can query additional ${category} orders within the next 2 minutes by simply typing the order numbers.\n\nType /menu for main menu or / to go back`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      
      userStates[from] = { 
        currentMenu: 'order_followup', 
        category: category,
        timestamp: Date.now()
      };
      
    } else {
      try {
        const searchTerm = orderNumbers.join(', ');
        const pdfResult = await generateOrderPDF(orderResults, searchTerm, from, category);
        
        const summaryMessage = `*Large Results Found*\n\nSearch: ${searchTerm}\nTotal Results: ${orderResults.length} orders\nPDF Generated: ${pdfResult.filename}\n\nResults are too many for WhatsApp\n\nType /menu for main menu`;
        
        await sendWhatsAppMessage(from, summaryMessage, productId, phoneId);
        await sendWhatsAppFile(from, pdfResult.filepath, pdfResult.filename, productId, phoneId, []);
        
        delete userStates[from];
        
      } catch (pdfError) {
        console.error('PDF generation failed:', pdfError);
        await sendWhatsAppMessage(from, `*Error Generating PDF*\n\nFound ${orderResults.length} results but could not generate PDF.\nPlease contact support.\n\nType /menu for main menu`, productId, phoneId);
        
        delete userStates[from];
      }
    }
    
  } catch (error) {
    console.error(`Error processing ${category} order query:`, error);
    await sendWhatsAppMessage(from, `Error checking ${category} orders\n\nPlease try again later.\n\nType /menu for main menu`, productId, phoneId);
    
    delete userStates[from];
  }
}

// MAIN WEBHOOK HANDLER
app.post('/webhook', async (req, res) => {
  const message = req.body.message?.text;
  const from = req.body.user?.phone;
  const productId = req.body.product_id || req.body.productId;
  const phoneId = req.body.phone_id || req.body.phoneId;

  if (!message || typeof message !== 'string') return res.sendStatus(200);

  const trimmedMessage = message.trim();
  if (trimmedMessage === '') return res.sendStatus(200);

  // Check for expired stock sessions
  if (userStates[from] && isStockSessionExpired(userStates[from])) {
    console.log(`Stock session expired for ${from} at ${new Date().toLocaleTimeString()}`);
    delete userStates[from];
  }

  // Check if this is a valid bot interaction
  if (!isValidBotInteraction(trimmedMessage, userStates[from])) {
    return res.sendStatus(200);
  }

  const lowerMessage = trimmedMessage.toLowerCase();

  // Debug commands
  if (lowerMessage === '/debuggreet') {
    const greeting = await getUserGreeting(from);
    const debugMessage = greeting ? `Found: ${greeting.salutation} ${greeting.name} - ${greeting.greetings}` : 'No greeting found in sheet';
    await sendWhatsAppMessage(from, `Greeting debug result: ${debugMessage}`, productId, phoneId);
    return res.sendStatus(200);
  }

  if (lowerMessage === '/debugpermissions') {
    const permissions = await getUserPermissions(from);
    const debugMessage = permissions.length > 0 ? `Your permissions: ${permissions.join(', ')}` : 'No permissions found';
    await sendWhatsAppMessage(from, `Permission debug result: ${debugMessage}`, productId, phoneId);
    return res.sendStatus(200);
  }

  if (lowerMessage.startsWith('/debugorder ')) {
    const testOrderNumber = trimmedMessage.replace('/debugorder ', '').trim();
    
    await sendWhatsAppMessage(from, `🔍 Debugging order search for: ${testOrderNumber}`, productId, phoneId);
    
    try {
      const auth = await getGoogleAuth();
      const authClient = await auth.getClient();
      const sheets = google.sheets({ version: 'v4', auth: authClient });
      
      const result = await searchInLiveSheet(sheets, testOrderNumber);
      
      const debugResult = result.found ? `✅ Order FOUND: ${result.message}` : `❌ Order NOT FOUND in live sheet`;
      await sendWhatsAppMessage(from, debugResult, productId, phoneId);
      
    } catch (error) {
      await sendWhatsAppMessage(from, `Error during debug: ${error.message}`, productId, phoneId);
    }
    
    return res.sendStatus(200);
  }

 // Enhanced debug command to find F1554O-1-2
if (lowerMessage.startsWith('/debugjacket ')) {
  const testOrderNumber = trimmedMessage.replace('/debugjacket ', '').trim();
  
  await sendWhatsAppMessage(from, `🔍 Debugging jacket search for: ${testOrderNumber}`, productId, phoneId);
  
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    
    // Get ALL data to find F1554O-1-2
    const testResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: JACKET_LIVE_SHEET_ID,
      range: `${JACKET_LIVE_SHEET_NAME}!A:F`, // Get ALL rows
      valueRenderOption: 'FORMATTED_VALUE'
    });
    
    const testRows = testResponse.data.values;
    let debugResult = `=== JACKET DEBUG RESULTS ===\n\n`;
    debugResult += `Search Term: "${testOrderNumber}"\n`;
    debugResult += `Total Rows Found: ${testRows.length}\n\n`;
    
    let foundMatches = [];
    let allOrderNumbers = [];
    
    if (testRows && testRows.length > 2) {
      // Check ALL rows for matches
      for (let i = 2; i < testRows.length; i++) {
        const row = testRows[i];
        if (row && row.length > 3 && row[3]) {
          const orderNum = row[3].toString().trim();
          if (orderNum !== '') {
            allOrderNumbers.push(`Row ${i}: ${orderNum}`);
            
            if (orderNum.toUpperCase().includes(testOrderNumber.toUpperCase())) {
              foundMatches.push(`Row ${i}: ${orderNum}`);
            }
          }
        }
      }
      
      debugResult += `All Order Numbers Found:\n`;
      allOrderNumbers.slice(0, 20).forEach(order => {
        debugResult += `${order}\n`;
      });
      
      debugResult += `\nMatches Found: ${foundMatches.length}\n`;
      if (foundMatches.length > 0) {
        foundMatches.forEach(match => {
          debugResult += `✅ ${match}\n`;
        });
      } else {
        debugResult += `❌ No orders contain "${testOrderNumber}"\n`;
        debugResult += `\nSearching specifically for "F1554O-1-2"...\n`;
        
        const f1554Match = allOrderNumbers.find(order => order.includes('F1554O-1-2'));
        if (f1554Match) {
          debugResult += `✅ Found F1554O-1-2: ${f1554Match}\n`;
        } else {
          debugResult += `❌ F1554O-1-2 not found in first ${allOrderNumbers.length} orders\n`;
        }
      }
    }
    
    await sendWhatsAppMessage(from, debugResult, productId, phoneId);
    
  } catch (error) {
    await sendWhatsAppMessage(from, `Error during jacket debug: ${error.message}`, productId, phoneId);
  }
  
  return res.sendStatus(200);
}
Key Changes:

✅ Searches ALL Rows: Not just first 10-20 rows

✅ Added Logging: Shows which rows it's checking

✅ Enhanced Debug: Looks specifically for F1554O-1-2

✅ Fixed Row Starting: Starts from row 2 (where data begins)

Test This:

Replace the jacket search function with the fixed version above

Add the enhanced debug command

Run /debugjacket 1554 - it should now find "F1554O-1-2"

Run /jacket then search "1554" - should find "F1554O-1-2"

The search for "1554" should now find "F1554O-1-2" since "F1554O-1-2" contains "1554"!

Related
Where in my sheet should I expect to find F1554O-1-2
Could hidden rows or filters hide the row with F1554O-1-2
How can I change my Sheets API range to ensure F1554O-1-2 is included
Why might the API return zero matches for F1554O-1-2 despite access
What debug steps can I run to confirm the row containing F1554O-1-2 exists




  if (lowerMessage === '/debugrows') {
    try {
      const auth = await getGoogleAuth();
      const authClient = await auth.getClient();
      const sheets = google.sheets({ version: 'v4', auth: authClient });
      
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: LIVE_SHEET_ID,
        range: `${LIVE_SHEET_NAME}!A1:E10`,
        valueRenderOption: 'UNFORMATTED_VALUE'
      });

      const rows = response.data.values;
      let debugMessage = `*LIVE SHEET DEBUG*\n\nFirst ${Math.min(10, rows.length)} rows:\n\n`;
      
      rows.forEach((row, index) => {
        const colD = row.length > 3 ? row[3] : 'EMPTY';
        debugMessage += `Row ${index}: ColD="${colD}"\n`;
      });
      
      await sendWhatsAppMessage(from, debugMessage, productId, phoneId);
      
    } catch (error) {
      await sendWhatsAppMessage(from, `Debug error: ${error.message}`, productId, phoneId);
    }
    
    return res.sendStatus(200);
  }

  // Handle single "/" for going back one step
  if (trimmedMessage === '/') {
    const wentBack = goBackOneStep(from);
    if (wentBack) {
      if (userStates[from] && userStates[from].currentMenu === 'main') {
        const greeting = await getUserGreeting(from);
        const personalizedMenu = await generatePersonalizedMenu(from);
        const finalMessage = formatGreetingMessage(greeting, personalizedMenu);
        await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
      } else if (userStates[from] && userStates[from].currentMenu === 'order_query') {
        const orderQueryMenu = `*ORDER QUERY*\n\nPlease select the product category:\n\n1. Shirting\n2. Jacket\n3. Trouser\n\nType the number to continue or / to go back`;
        await sendWhatsAppMessage(from, orderQueryMenu, productId, phoneId);
      }
    } else {
      await sendWhatsAppMessage(from, 'Cannot go back further. Type /menu for main menu.', productId, phoneId);
    }
    return res.sendStatus(200);
  }

  // CONTEXT SWITCHING SHORTCUTS - These can switch from ANY state immediately

  // Main menu
  if (lowerMessage === '/menu') {
    userStates[from] = { currentMenu: 'main', timestamp: Date.now() };
    
    const greeting = await getUserGreeting(from);
    const personalizedMenu = await generatePersonalizedMenu(from);
    
    const finalMessage = formatGreetingMessage(greeting, personalizedMenu);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Direct Stock Query shortcut with 40-second session
  if (lowerMessage === '/stock') {
    if (!(await hasFeatureAccess(from, 'stock'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Stock Query.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    userStates[from] = { 
      currentMenu: 'smart_stock_query', 
      lastActivity: Date.now() // Start 40-second session
    };
    
    const greeting = await getUserGreeting(from);
    
    const stockQueryPrompt = `*SMART STOCK QUERY*\n\nEnter any 5+ character code (letters/numbers):\n\nExamples:\n- 11010 (finds 11010088471-001)\n- ABC12 (finds ABC123456-XYZ)\n- 88471 (finds 11010088471-001)\n\nMultiple searches: Separate with commas\nExample: 11010, ABC12, 88471\n\nSmart search finds partial matches\n\nType your search terms below or / to go back:`;
    
    const finalMessage = formatGreetingMessage(greeting, stockQueryPrompt);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Direct Order Query shortcut
  if (lowerMessage === '/order') {
    if (!(await hasFeatureAccess(from, 'order'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Order Query.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    userStates[from] = { currentMenu: 'order_query', timestamp: Date.now() };
    
    const greeting = await getUserGreeting(from);
    
    const orderQueryMenu = `*ORDER QUERY*\n\nPlease select the product category:\n\n1. Shirting\n2. Jacket\n3. Trouser\n\nType the number to continue or / to go back`;
    
    const finalMessage = formatGreetingMessage(greeting, orderQueryMenu);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Order category shortcuts with security check
  if (lowerMessage === '/shirting' || lowerMessage === '/jacket' || lowerMessage === '/trouser') {
    if (!(await hasFeatureAccess(from, 'order'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Order Query.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    const categoryMap = {
      '/shirting': 'Shirting',
      '/jacket': 'Jacket', 
      '/trouser': 'Trouser'
    };
    
    const category = categoryMap[lowerMessage];
    userStates[from] = { currentMenu: 'order_number_input', category: category, timestamp: Date.now() };
    
    const greeting = await getUserGreeting(from);
    
    const orderQuery = `*${category.toUpperCase()} ORDER QUERY*\n\nPlease enter your Order Number(s) or search terms:\n\nFull order: GT54695O-1-1, D47727S-1-2\nPartial search: GT546, D477, 1554\nMultiple: GT546, D477, 1554\n\nType your search terms below or / to go back:`;
    
    const finalMessage = formatGreetingMessage(greeting, orderQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Direct ticket shortcuts
  if (lowerMessage === '/helpticket') {
    if (!(await hasFeatureAccess(from, 'help_ticket'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Help Ticket.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    const helpTicketMessage = `*HELP TICKET*\n\nClick the link below to access Help Ticket form directly:\n\n${links.helpTicket}\n\nType /menu for main menu or / to go back`;

    await sendWhatsAppMessage(from, helpTicketMessage, productId, phoneId);
    delete userStates[from]; // Clear any existing state
    return res.sendStatus(200);
  }

  if (lowerMessage === '/delegation') {
    if (!(await hasFeatureAccess(from, 'delegation'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Delegation.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    const delegationMessage = `*DELEGATION*\n\nClick the link below to access Delegation form directly:\n\n${links.delegation}\n\nType /menu for main menu or / to go back`;

    await sendWhatsAppMessage(from, delegationMessage, productId, phoneId);
    delete userStates[from]; // Clear any existing state
    return res.sendStatus(200);
  }

  // Handle order follow-up queries within 2-minute window
  if (userStates[from] && userStates[from].currentMenu === 'order_followup') {
    if (isWithinOrderQueryWindow(from) && trimmedMessage !== '/menu' && trimmedMessage !== '/') {
      const orderNumbers = trimmedMessage.split(',').map(order => order.trim()).filter(order => order.length > 0);
      
      if (orderNumbers.length > 0) {
        await processOrderQuery(from, userStates[from].category, orderNumbers, productId, phoneId, true);
        return res.sendStatus(200);
      }
    } else {
      userStates[from] = { currentMenu: 'main', timestamp: Date.now() };
    }
  }

  // Handle menu selections with fine-grained security checks
  if (userStates[from] && userStates[from].currentMenu === 'main') {
    
    // TICKET OPTION (1)
    if (trimmedMessage === '1') {
      const hasAnyTicketAccess = (await hasFeatureAccess(from, 'help_ticket')) ||
                                 (await hasFeatureAccess(from, 'delegation')) ||
                                 (await hasFeatureAccess(from, 'leave_form'));
      
      if (!hasAnyTicketAccess) {
        await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Ticket options.\nContact administrator for access.`, productId, phoneId);
        return res.sendStatus(200);
      }

      const ticketMenu = await generateTicketMenu(from);
      await sendWhatsAppMessage(from, ticketMenu, productId, phoneId);
      delete userStates[from]; // Clear state after showing tickets
      return res.sendStatus(200);
    }

    // ORDER QUERY OPTION (2)
    if (trimmedMessage === '2' && (await hasFeatureAccess(from, 'order'))) {
      userStates[from] = { currentMenu: 'order_query', timestamp: Date.now() };
      const orderQueryMenu = `*ORDER QUERY*\n\nPlease select the product category:\n\n1. Shirting\n2. Jacket\n3. Trouser\n\nType the number to continue or / to go back`;
      
      await sendWhatsAppMessage(from, orderQueryMenu, productId, phoneId);
      return res.sendStatus(200);
    }

    // STOCK QUERY OPTION (3)
    if (trimmedMessage === '3' && (await hasFeatureAccess(from, 'stock'))) {
      userStates[from] = { 
        currentMenu: 'smart_stock_query', 
        lastActivity: Date.now() // Start 40-second session
      };
      
      const greeting = await getUserGreeting(from);
      
      const stockQueryPrompt = `*SMART STOCK QUERY*\n\nEnter any 5+ character code (letters/numbers):\n\nExamples:\n- 11010 (finds 11010088471-001)\n- ABC12 (finds ABC123456-XYZ)\n- 88471 (finds 11010088471-001)\n\nMultiple searches: Separate with commas\nExample: 11010, ABC12, 88471\n\nSmart search finds partial matches\n\nType your search terms below or / to go back:`;
      
      const finalMessage = formatGreetingMessage(greeting, stockQueryPrompt);
      await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
      return res.sendStatus(200);
    }

    // DOCUMENT OPTION (4)
    if (trimmedMessage === '4' && (await hasFeatureAccess(from, 'document'))) {
      await sendWhatsAppMessage(from, '*DOCUMENT*\n\nThis feature is coming soon!\n\nType /menu for main menu or / to go back', productId, phoneId);
      delete userStates[from]; // Clear state
      return res.sendStatus(200);
    }

    // If user tries to access unauthorized feature
    await sendWhatsAppMessage(from, '*ACCESS DENIED*\n\nYou don\'t have permission for this option or invalid selection.\n\nType /menu to see your available options.', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle order query category selection
  if (userStates[from] && userStates[from].currentMenu === 'order_query') {
    if (trimmedMessage === '1') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Shirting', timestamp: Date.now() };
      await sendWhatsAppMessage(from, `*SHIRTING ORDER QUERY*\n\nPlease enter your Order Number(s) or search terms:\n\nFull order: B-J3005Z-1-1\nPartial search: J3005Z, J300, 1234\nMultiple: J3005Z, J300, 1234\n\nType your search terms below or / to go back:`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Jacket', timestamp: Date.now() };
      await sendWhatsAppMessage(from, `*JACKET ORDER QUERY*\n\nPlease enter your Order Number(s) or search terms:\n\nFull order: GT54695O-1-1, D47727S-1-2\nPartial search: GT546, D477, 1554\nMultiple: GT546, D477, 1554\n\nType your search terms below or / to go back:`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Trouser', timestamp: Date.now() };
      await sendWhatsAppMessage(from, `*TROUSER ORDER QUERY*\n\nPlease enter your Order Number(s) or search terms:\n\nFull order: TR54695O-1-1\nPartial search: TR546, 1234\nMultiple: TR546, 1234\n\nType your search terms below or / to go back:`, productId, phoneId);
      return res.sendStatus(200);
    }

    await sendWhatsAppMessage(from, 'Invalid option. Please select 1, 2, or 3.\n\nType /menu for main menu or / to go back', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle order number input
  if (userStates[from] && userStates[from].currentMenu === 'order_number_input') {
    if (trimmedMessage !== '/menu' && trimmedMessage !== '/') {
      const category = userStates[from].category;
      const orderNumbers = trimmedMessage.split(',').map(order => order.trim()).filter(order => order.length > 0);
      
      await processOrderQuery(from, category, orderNumbers, productId, phoneId);
      return res.sendStatus(200);
    }
  }

  // Handle smart stock query input with 40-second session management
  if (userStates[from] && userStates[from].currentMenu === 'smart_stock_query') {
    if (trimmedMessage !== '/menu' && trimmedMessage !== '/') {
      const searchTerms = trimmedMessage.split(',').map(q => q.trim()).filter(q => q.length > 0);
      await processSmartStockQuery(from, searchTerms, productId, phoneId);
      return res.sendStatus(200);
    }
  }

  return res.sendStatus(200);
});

async function sendWhatsAppMessage(to, message, productId, phoneId) {
  try {
    await axios.post(
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
  } catch (error) {
    console.error('Error sending message:', error.response?.data || error.message);
  }
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`WhatsApp Bot running on port ${PORT}`);
  console.log('✅ FIXED: All context switching shortcuts work from any state');
  console.log('✅ FIXED: 40-second session timer resets with each stock query');
  console.log('✅ FIXED: Commands can switch contexts immediately');
  console.log('✅ FIXED: Bot ignores casual messages appropriately');
  console.log('✅ NEW: Enhanced order search with flexible pattern matching');
  console.log('✅ NEW: Smart matching for J3005Z, J300 → B-J3005Z-1-1, B-J3005Y-1-2, etc.');
  console.log('✅ NEW: JACKET WORKFLOW INTEGRATED - Separate search system');
  console.log('✅ FIXED: Enhanced jacket order search with better error handling');
  console.log('✅ FIXED: 1554 matching issue resolved with flexible matching');
  console.log('✅ All existing functions remain intact');
  console.log('');
  console.log('🚀 CONTEXT SWITCHING SHORTCUTS:');
  console.log('   /menu - Main menu (from anywhere)');
  console.log('   /stock - Direct stock query (from anywhere)');
  console.log('   /order - Order query menu (from anywhere)');
  console.log('   /shirting - Direct shirting orders (from anywhere)');
  console.log('   /jacket - Direct jacket orders (from anywhere)'); 
  console.log('   /trouser - Direct trouser orders (from anywhere)');
  console.log('   /helpticket - Direct help ticket (from anywhere)');
  console.log('   /delegation - Direct delegation (from anywhere)');
  console.log('');
  console.log('🔍 ENHANCED ORDER SEARCH FEATURES:');
  console.log('   - Full order numbers: GT54695O-1-1, D47727S-1-2');
  console.log('   - Partial codes: GT546 → finds GT54695O-1-1');
  console.log('   - Category-specific: Shirting and Jacket separate systems');
  console.log('   - Pattern matching with dashes, spaces, underscores');
  console.log('   - PDF generation for >3 results');
  console.log('   - WhatsApp display for ≤3 results');
  console.log('');
  console.log('📋 JACKET PRODUCTION STAGES:');
  console.log('   CUT → FUS → Prep → MAK → QC1 → BH → Press → QC2 → Dispatch(Factory) → Dispatch(HO)');
  console.log('');
  console.log('Debug commands: /debuggreet, /debugpermissions, /debugorder, /debugjacket, /debugrows');
});
