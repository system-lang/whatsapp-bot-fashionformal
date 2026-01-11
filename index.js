require('dotenv').config();
const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.json());

// File serving endpoint for PDFs
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

// Store user states and order query timestamps
let userStates = {};
let orderQueryTimestamps = {};

// Configuration constants
const STOCK_SESSION_TIMEOUT = 40 * 1000;
const SHORTCUT_COMMANDS = [
  '/menu', '/stock', '/order', '/shirting', '/jacket', '/trouser',
  '/helpticket', '/delegation', '/',
  '/debuggreet', '/debugpermissions', '/debugrows', '/debugjacket', '/debugtrouser'
];

const links = {
  helpTicket: 'https://tinyurl.com/HelpticketFF',
  delegation: 'https://tinyurl.com/ffdelegation1',
  leave: 'YOUR_LEAVE_FORM_LINK_HERE'
};

const MAYTAPI_API_TOKEN = '07d75e68-b94f-485b-9e8c-19e707d176ae';

// Google Sheets configuration
const STOCK_FOLDER_ID = '1QV1cJ9jJZZW2PY24uUY2hefKeUqVHrrf';
const STORE_PERMISSION_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';
const GREETINGS_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';
const USER_ACCESS_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';
const STATIC_FORM_BASE_URL = 'https://docs.google.com/forms/d/e/1FAIpQLSfyAo7LwYtQDfNVxPRbHdk_ymGpDs-RyWTCgzd2PdRhj0T3Hw/viewform';

// Order configurations
const LIVE_SHEET_ID = '1AxjCHsMxYUmEULaW1LxkW78g0Bv9fp4PkZteJO82uEA';
const LIVE_SHEET_NAME = 'FMS';
const COMPLETED_ORDER_FOLDER_ID = '1kgdPdnUK-FsnKZDE5yW6vtRf2H9d3YRE';

const JACKET_LIVE_SHEET_ID = '1XYXOv6C-aIuMVYDLSMflPZIQL7yJWq5BgnmDAnRMt58';
const JACKET_LIVE_SHEET_NAME = 'FMS';
const JACKET_COMPLETED_ORDER_FOLDER_ID = '1GmcGommmEBlP4iNPRA6NC4nbyFokCdY8';

// NEW: Trouser Order Configuration
const TROUSER_LIVE_SHEET_ID = '1y96TQMTrWXgAQmXqiXtcj3WGtdmcuMJS4F7a8OU_Sk4';
const TROUSER_LIVE_SHEET_NAME = 'FMS';
const TROUSER_COMPLETED_ORDER_FOLDER_ID = '104EOy6nU35CwZ_vlpjaCywC7yjlovIZ-';

// Production stages configurations
const PRODUCTION_STAGES = [
  { name: 'CUT', column: 'T' },
  { name: 'FUS', column: 'Z' },
  { name: 'PAS', column: 'AF' },
  { name: 'MAK', column: 'AL' },
  { name: 'BH',  column: 'AX' },
  { name: 'BS',  column: 'BD' },
  { name: 'QC',  column: 'BJ' },
  { name: 'ALT', column: 'BP' },
  { name: 'IRO', column: 'BY' },
  { name: 'Dispatch (Factory)', column: 'CE' },
  { name: 'Dispatch (HO)',      column: 'CL', dispatchDateColumn: 'CL' }
];

const JACKET_PRODUCTION_STAGES = [
  { name: 'CUT', column: 'T' },
  { name: 'FUS', column: 'Z' },
  { name: 'Prep', column: 'AF' },
  { name: 'MAK', column: 'AL' },
  { name: 'QC1', column: 'AR' },
  { name: 'BH',  column: 'AX' },
  { name: 'Press', column: 'BD' },
  { name: 'QC2', column: 'BJ' },
  { name: 'Dispatch (Factory)', column: 'BP' },
  { name: 'Dispatch (HO)',      column: 'BW', dispatchDateColumn: 'BW' }
];

// NEW: Trouser Production stages (same as jacket)
const TROUSER_PRODUCTION_STAGES = [
  { name: 'CUT', column: 'T' },
  { name: 'FUS', column: 'Z' },
  { name: 'Prep', column: 'AF' },
  { name: 'MAK', column: 'AL' },
  { name: 'QC1', column: 'AR' },
  { name: 'BH',  column: 'AX' },
  { name: 'Press', column: 'BD' },
  { name: 'QC2', column: 'BJ' },
  { name: 'Dispatch (Factory)', column: 'BP' },
  { name: 'Dispatch (HO)',      column: 'BW', dispatchDateColumn: 'BW' }
];


// Utility Functions
function isStockSessionExpired(userState) {
  if (!userState || userState.currentMenu !== 'smart_stock_query' || !userState.lastActivity) {
    return false;
  }
  const now = Date.now();
  const timeSinceLastActivity = now - userState.lastActivity;
  return timeSinceLastActivity > STOCK_SESSION_TIMEOUT;
}

function updateLastActivity(from) {
  if (userStates[from] && userStates[from].currentMenu === 'smart_stock_query') {
    userStates[from].lastActivity = Date.now();
    console.log(`Updated stock session activity for ${from} at ${new Date().toLocaleTimeString()}`);
  }
}

function isValidBotInteraction(message, userState) {
  const lowerMessage = message.toLowerCase().trim();

  if (lowerMessage.startsWith('/debugorder ')) {
    return true;
  }

  if (SHORTCUT_COMMANDS.includes(lowerMessage)) {
    return true;
  }

  if (['1', '2', '3', '4'].includes(message.trim())) {
    return true;
  }

  if (userState) {
    switch (userState.currentMenu) {
      case 'main':
        return ['1', '2', '3', '4'].includes(message.trim());
      case 'order_number_input': // MODIFIED: Remove order_query state check
      case 'order_followup':
        return message.trim().length > 0;
      case 'smart_stock_query':
        if (isStockSessionExpired(userState)) {
          return false;
        }
        return message.trim().length > 0;
      default:
        return false;
    }
  }

  return false;
}

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

function formatDateForDisplay(rawDate) {
  if (!rawDate || rawDate === '') {
    return 'Date not available';
  }

  const dateStr = rawDate.toString().trim();

  if (dateStr.includes('/') || dateStr.includes('-') || dateStr.includes(' ')) {
    return dateStr;
  }

  const dateNum = parseFloat(dateStr);
  if (!isNaN(dateNum) && dateNum > 1000) {
    try {
      const jsDate = new Date((dateNum - 25569) * 86400 * 1000);

      if (!isNaN(jsDate.getTime())) {
        return jsDate.toLocaleString('en-IN', { 
          timeZone: 'Asia/Kolkata',
          year: 'numeric',
          month: '2-digit', 
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit'
        });
      }
    } catch (error) {
      // If conversion fails, return original
    }
  }

  return dateStr;
}

function formatDateWithSuffix(rawDate) {
  if (!rawDate || rawDate === '') {
    return 'Date not available';
  }

  let jsDate = null;
  const dateStr = rawDate.toString().trim();

  // Try normal date string
  if (dateStr.includes('/') || dateStr.includes('-') || dateStr.includes(' ')) {
    const parsed = new Date(dateStr);
    if (!isNaN(parsed.getTime())) {
      jsDate = parsed;
    }
  } else {
    // Try Google Sheets serial number
    const dateNum = parseFloat(dateStr);
    if (!isNaN(dateNum) && dateNum > 1000) {
      jsDate = new Date((dateNum - 25569) * 86400 * 1000);
    }
  }

  if (!jsDate || isNaN(jsDate.getTime())) {
    return dateStr;
  }

  const day = jsDate.getDate();
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = monthNames[jsDate.getMonth()];
  const year = jsDate.getFullYear();

  const j = day % 10, k = day % 100;
  let suffix = 'th';
  if (j === 1 && k !== 11) suffix = 'st';
  else if (j === 2 && k !== 12) suffix = 'nd';
  else if (j === 3 && k !== 13) suffix = 'rd';

  return `${day}${suffix} ${month} ${year}`;
}


function formatStockQuantity(stockValue) {
  if (!stockValue || stockValue === '') return stockValue;

  const numValue = parseFloat(stockValue.toString().trim());
  if (!isNaN(numValue) && numValue > 15) {
    return '15+';
  }

  return stockValue.toString();
}

// MODIFIED: Updated goBackOneStep to handle new direct order flow
function goBackOneStep(from) {
  if (!userStates[from]) return false;

  const currentMenu = userStates[from].currentMenu;

  // REMOVED: order_query state handling since we skip that menu
  if (currentMenu === 'order_number_input') {
    userStates[from] = { currentMenu: 'main', timestamp: Date.now() };
    return true;
  }

  if (currentMenu === 'smart_stock_query') {
    userStates[from] = { currentMenu: 'main', timestamp: Date.now() };
    return true;
  }

  if (currentMenu === 'main' || currentMenu === 'completed') {
    return false;
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

function isOrderMatch(orderNumber, searchTerm) {
  const upperOrderNumber = orderNumber.toUpperCase();
  const upperSearchTerm = searchTerm.toUpperCase();

  const cleanOrderNumber = upperOrderNumber.replace(/[-_\s]/g, '');
  const cleanSearchTerm = upperSearchTerm.replace(/[-_\s]/g, '');

  if (upperOrderNumber.includes(upperSearchTerm)) {
    return true;
  }

  if (cleanOrderNumber.includes(cleanSearchTerm)) {
    return true;
  }

  const orderSegments = upperOrderNumber.split(/[-_\s]/);
  for (const segment of orderSegments) {
    if (segment.includes(upperSearchTerm)) {
      return true;
    }
  }

  if (upperSearchTerm.length >= 3) {
    const regex = new RegExp(upperSearchTerm.split('').join('.*?'), 'i');
    if (regex.test(upperOrderNumber)) {
      return true;
    }
  }

  return false;
}

function isWithinOrderQueryWindow(from) {
  if (!orderQueryTimestamps[from]) return false;

  const now = Date.now();
  const lastQuery = orderQueryTimestamps[from];
  const twoMinutes = 2 * 60 * 1000;

  return (now - lastQuery) < twoMinutes;  
}
function checkProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let hasAnyStage = false;
    let lastStageDate = '';

    for (let i = 0; i < PRODUCTION_STAGES.length; i++) {
      const stage = PRODUCTION_STAGES[i];
      const columnIndex = columnToIndex(stage.column);

let cellValue = '';
      if (row.length > columnIndex && row[columnIndex] !== undefined && row[columnIndex] !== null) {
        cellValue = row[columnIndex].toString().trim();
      }

  if (cellValue !== '') {
 lastCompletedStage = stage;
        lastStageDate = cellValue;
        hasAnyStage = true;
      }
    }

    if (!hasAnyStage) {
      return { message: 'Order is currently under process' };
    }

    if (lastCompletedStage && lastCompletedStage.name === 'Dispatch (HO)') {
      const dispatchDateIndex = columnToIndex(lastCompletedStage.dispatchDateColumn);

      
      let rawDispatchDate = '';
      if (row.length > dispatchDateIndex && row[dispatchDateIndex] !== undefined && row[dispatchDateIndex] !== null) {
        rawDispatchDate = row[dispatchDateIndex];
      }

 const formattedDate = formatDateWithSuffix(rawDispatchDate);
return { message: `Order has been dispatched from HO on ${formattedDate}` };
    }

    const formattedStageDate = formatDateWithSuffix(lastStageDate);
    return { message: `Order is completed ${lastCompletedStage.name} on ${formattedStageDate}` };
 } catch (error) {
    console.error('Error checking production stages:', error);
    return { message: 'Error checking order status' };
  }
}
function checkJacketProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let hasAnyStage = false;
    let lastStageDate = '';

    for (let i = 0; i < JACKET_PRODUCTION_STAGES.length; i++) {
      const stage = JACKET_PRODUCTION_STAGES[i];
      const columnIndex = columnToIndex(stage.column);
 let cellValue = '';
      if (row.length > columnIndex && row[columnIndex] !== undefined && row[columnIndex] !== null) {
        cellValue = row[columnIndex].toString().trim();
      }

      if (cellValue !== '') {
        lastCompletedStage = stage;
        lastStageDate = cellValue;
        hasAnyStage = true;
      }
    }

    if (!hasAnyStage) {
      return { message: 'Order is currently under process' };
    }

    if (lastCompletedStage && lastCompletedStage.name === 'Dispatch (HO)') {
      const dispatchDateIndex = columnToIndex(lastCompletedStage.dispatchDateColumn);

       let rawDispatchDate = '';
      if (row.length > dispatchDateIndex && row[dispatchDateIndex] !== undefined && row[dispatchDateIndex] !== null) {
        rawDispatchDate = row[dispatchDateIndex];
      }

      const formattedDate = formatDateWithSuffix(rawDispatchDate);
return { message: `Order has been dispatched from HO on ${formattedDate}` };
    }
 const formattedStageDate = formatDateWithSuffix(lastStageDate);
    return { message: `Order is completed ${lastCompletedStage.name} on ${formattedStageDate}` };
} catch (error) {
    console.error('Error checking jacket production stages:', error);
    return { message: 'Error checking order status' };
  }
}
function checkTrouserProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let hasAnyStage = false;
    let lastStageDate = '';

    for (let i = 0; i < TROUSER_PRODUCTION_STAGES.length; i++) {
      const stage = TROUSER_PRODUCTION_STAGES[i];
      const columnIndex = columnToIndex(stage.column);

      
      let cellValue = '';
      if (row.length > columnIndex && row[columnIndex] !== undefined && row[columnIndex] !== null) {
        cellValue = row[columnIndex].toString().trim();
      }

      if (cellValue !== '') {
 lastCompletedStage = stage;
        lastStageDate = cellValue;
        hasAnyStage = true;
      }
    }

    if (!hasAnyStage) {
      return { message: 'Order is currently under process' };
    }

    if (lastCompletedStage && lastCompletedStage.name === 'Dispatch (HO)') {
      const dispatchDateIndex = columnToIndex(lastCompletedStage.dispatchDateColumn);

      
      let rawDispatchDate = '';
      if (row.length > dispatchDateIndex && row[dispatchDateIndex] !== undefined && row[dispatchDateIndex] !== null) {
        rawDispatchDate = row[dispatchDateIndex];
      }

 const formattedDate = formatDateWithSuffix(rawDispatchDate);
return { message: `Order has been dispatched from HO on ${formattedDate}` };
    }

    const formattedStageDate = formatDateWithSuffix(lastStageDate);
    return { message: `Order is completed ${lastCompletedStage.name} on ${formattedStageDate}` };
} catch (error) {
    console.error('Error checking trouser production stages:', error);
    return { message: 'Error checking order status' };
  }
}

// User Permission Functions (UNCHANGED)
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
    if (!rows || rows.length === 0) {
      return [];
    }

    const phoneVariations = [
      phoneNumber,
      phoneNumber.replace(/^\+91/, ''),
      phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''),
      phoneNumber.replace(/^0/, ''),
      phoneNumber.slice(-10)
    ];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 2) continue;

      const sheetPhone = (row[0] || '').toString().trim();
      const featuresString = (row[1] || '').toString().trim();

      for (const phoneVar of phoneVariations) {
        if (phoneVar === sheetPhone) {
          const features = featuresString.split(',').map(f => f.trim().toLowerCase());
          return features;
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

// Menu Generation Functions (UNCHANGED)
async function generatePersonalizedMenu(phoneNumber) {
  const userPermissions = await getUserPermissions(phoneNumber);

  if (userPermissions.length === 0) {
    return `*ACCESS DENIED*

You do not have permission to use this bot.
Please contact administrator for access.`;
  }

  let menuItems = [];
  let shortcuts = [];

  const hasAnyTicketAccess = userPermissions.some(perm => 
    ['help_ticket', 'delegation', 'leave_form'].includes(perm)
  );

  if (hasAnyTicketAccess) {
    menuItems.push('1. Ticket');

    if (userPermissions.includes('help_ticket')) {
      shortcuts.push('/helpticket - Direct Help Ticket');
    }
    if (userPermissions.includes('delegation')) {
      shortcuts.push('/delegation - Direct Delegation');
    }
  }

  if (userPermissions.includes('order')) {
    menuItems.push('2. Order Query');
    shortcuts.push('/order - Direct Order Search'); // MODIFIED: Updated text
    shortcuts.push('/shirting - Shirting Orders');
    shortcuts.push('/jacket - Jacket Orders');
    shortcuts.push('/trouser - Trouser Orders');
  }

  if (userPermissions.includes('stock')) {
    menuItems.push('3. Stock Query');
    shortcuts.push('/stock - Direct Stock Query');
  }

  if (userPermissions.includes('document')) {
    menuItems.push('4. Document');
  }

  if (menuItems.length === 0) {
    return `*ACCESS DENIED*

No valid permissions found.
Please contact administrator for access.`;
  }

  let menu = `*MAIN MENU*

Please select an option:

${menuItems.join('\n')}`;

  if (shortcuts.length > 0) {
    menu += `\n\n*SHORTCUTS:*\n${shortcuts.join('\n')}`;
  }

  menu += `\n\nType the number, use shortcuts, or / to go back`;

  return menu;
}

async function generateTicketMenu(phoneNumber) {
  const userPermissions = await getUserPermissions(phoneNumber);

  let ticketOptions = [];

  if (userPermissions.includes('help_ticket')) {
    ticketOptions.push(`*HELP TICKET*\n${links.helpTicket}`);
  }

  if (userPermissions.includes('delegation')) {
    ticketOptions.push(`*DELEGATION*\n${links.delegation}`);
  }

  if (userPermissions.includes('leave_form')) {
    ticketOptions.push(`*LEAVE FORM*\n${links.leave}`);
  }

  if (ticketOptions.length === 0) {
    return `*ACCESS DENIED*

You don't have permission to access any ticket options.
Contact administrator for access.

Type /menu to return to main menu`;
  }

  return `*TICKET OPTIONS*

Click the links below to access forms directly:

${ticketOptions.join('\n\n')}

Type /menu to return to main menu or / to go back`;
}

async function getUserGreeting(phoneNumber) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const attempts = [
      'Greetings!A:D',
      "'Greetings'!A:D", 
      'Sheet2!A:D',
      'A:D'
    ];

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
      } catch (rangeError) {
        continue;
      }
    }

    if (!rows || rows.length <= 1) {
      return null;
    }

    const phoneVariations = [
      phoneNumber,
      phoneNumber.replace(/^\+91/, ''),
      phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''),
      phoneNumber.replace(/^0/, ''),
      phoneNumber.slice(-10)
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
  if (!greeting || !greeting.name || !greeting.salutation || !greeting.greetings) {
    return mainMessage;
  }

  return `${greeting.salutation} ${greeting.name}\n\n${greeting.greetings}\n\n${mainMessage}`;
}

// NEW: UNIFIED ORDER SEARCH FUNCTIONS - Super Fast Parallel Search
async function searchAllOrdersUnified(searchTerm) {
  const matchingOrders = [];

  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    const cleanSearchTerm = searchTerm.trim().toUpperCase();

    console.log(`🔍 UNIFIED SEARCH for: "${cleanSearchTerm}"`);

    // PARALLEL SEARCH: Search all 3 product types simultaneously using Promise.all()
    const [shirtingResults, jacketResults, trouserResults] = await Promise.all([
      searchShirtingOrders(cleanSearchTerm, auth, authClient, sheets, drive),
      searchJacketOrders(cleanSearchTerm, auth, authClient, sheets, drive),
      searchTrouserOrders(cleanSearchTerm, auth, authClient, sheets, drive)
    ]);

    // Combine results from all product types
    matchingOrders.push(...shirtingResults);
    matchingOrders.push(...jacketResults);
    matchingOrders.push(...trouserResults);

    console.log(`📊 UNIFIED SEARCH completed: ${matchingOrders.length} total matches`);
    return matchingOrders;

  } catch (error) {
    console.error('Error in searchAllOrdersUnified:', error);
    return [];
  }
}

// Helper function for Shirting search
async function searchShirtingOrders(cleanSearchTerm, auth, authClient, sheets, drive) {
  const results = [];

  try {
    // Search shirting live sheet
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
          results.push({
            orderNumber: orderNumber,
            message: stageStatus.message,
            location: 'Live Sheet',
            productType: 'Shirting'
          });
        }
      }
    }

    // Search shirting completed orders
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

            results.push({
              orderNumber: orderNumber,
              message: `Order got dispatched on ${formattedDate}`,
              location: 'Completed Orders',
              productType: 'Shirting'
            });
          }
        }
      } catch (error) {
        continue;
      }
    }

  } catch (error) {
    console.error('Error searching shirting orders:', error);
  }

  return results;
}

// Helper function for Jacket search
async function searchJacketOrders(cleanSearchTerm, auth, authClient, sheets, drive) {
  const results = [];

  try {
    // Search jacket live sheet
    const liveResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: JACKET_LIVE_SHEET_ID,
      range: `${JACKET_LIVE_SHEET_NAME}!A:BW`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const liveRows = liveResponse.data.values;
    if (liveRows && liveRows.length > 2) {
      for (let i = 2; i < liveRows.length; i++) {
        const row = liveRows[i];
        if (!row || row.length === 0) continue;

        let orderNumber = '';
        if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
          orderNumber = row[3].toString().trim();
        }

        if (orderNumber && isOrderMatch(orderNumber, cleanSearchTerm)) {
          const stageStatus = checkJacketProductionStages(row);
          results.push({
            orderNumber: orderNumber,
            message: stageStatus.message,
            location: 'Live Sheet',
            productType: 'Jacket'
          });
        }
      }
    }

    // Search jacket completed orders
    const folderFiles = await drive.files.list({
      q: `'${JACKET_COMPLETED_ORDER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    for (const file of folderFiles.data.files) {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: file.id,
          range: 'A:BW',
          valueRenderOption: 'UNFORMATTED_VALUE'
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) continue;

        for (let i = 2; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          let orderNumber = '';
          if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
            orderNumber = row[3].toString().trim();
          }

          if (orderNumber && isOrderMatch(orderNumber, cleanSearchTerm)) {
            let rawDispatchDate = '';
            if (row.length > 74 && row[74] !== undefined && row[74] !== null) {
              rawDispatchDate = row[74];
            }

            const formattedDate = formatDateForDisplay(rawDispatchDate);

            results.push({
              orderNumber: orderNumber,
              message: `Order got dispatched on ${formattedDate}`,
              location: 'Completed Orders',
              productType: 'Jacket'
            });
          }
        }
      } catch (error) {
        continue;
      }
    }

  } catch (error) {
    console.error('Error searching jacket orders:', error);
  }

  return results;
}

// Helper function for Trouser search
async function searchTrouserOrders(cleanSearchTerm, auth, authClient, sheets, drive) {
  const results = [];

  try {
    // Search trouser live sheet
    const liveResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: TROUSER_LIVE_SHEET_ID,
      range: `${TROUSER_LIVE_SHEET_NAME}!A:BW`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const liveRows = liveResponse.data.values;
    if (liveRows && liveRows.length > 2) {
      for (let i = 2; i < liveRows.length; i++) {
        const row = liveRows[i];
        if (!row || row.length === 0) continue;

        let orderNumber = '';
        if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
          orderNumber = row[3].toString().trim();
        }

        if (orderNumber && isOrderMatch(orderNumber, cleanSearchTerm)) {
          const stageStatus = checkTrouserProductionStages(row);
          results.push({
            orderNumber: orderNumber,
            message: stageStatus.message,
            location: 'Live Sheet',
            productType: 'Trouser'
          });
        }
      }
    }

    // Search trouser completed orders
    const folderFiles = await drive.files.list({
      q: `'${TROUSER_COMPLETED_ORDER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    for (const file of folderFiles.data.files) {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: file.id,
          range: 'A:BW',
          valueRenderOption: 'UNFORMATTED_VALUE'
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) continue;

        for (let i = 2; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          let orderNumber = '';
          if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
            orderNumber = row[3].toString().trim();
          }

          if (orderNumber && isOrderMatch(orderNumber, cleanSearchTerm)) {
            let rawDispatchDate = '';
            if (row.length > 74 && row[74] !== undefined && row[74] !== null) {
              rawDispatchDate = row[74];
            }

            const formattedDate = formatDateForDisplay(rawDispatchDate);

            results.push({
              orderNumber: orderNumber,
              message: `Order got dispatched on ${formattedDate}`,
              location: 'Completed Orders',
              productType: 'Trouser'
            });
          }
        }
      } catch (error) {
        continue;
      }
    }

  } catch (error) {
    console.error('Error searching trouser orders:', error);
  }

  return results;
}

// MODIFIED: New unified processOrderQuery function
async function processOrderQuery(from, orderNumbers, productId, phoneId, isFollowUp = false) {
  try {
    if (!isFollowUp) {
      await sendWhatsAppMessage(from, `*SEARCHING ALL ORDERS...*

🔍 Checking Shirting orders...
🔍 Checking Jacket orders...
🔍 Checking Trouser orders...

Please wait while I search across all product types.`, productId, phoneId);
    }

    let allResults = [];

    // Process each order number/search term
    for (const orderInput of orderNumbers) {
      const cleanInput = orderInput.trim();

      console.log(`Processing unified search for: "${cleanInput}"`);

      // Search across all product types simultaneously
      const matchingOrders = await searchAllOrdersUnified(cleanInput);
      allResults = allResults.concat(matchingOrders);
    }

    // Remove duplicates
    const uniqueResults = [];
    const seenOrders = new Set();

    for (const result of allResults) {
      const orderKey = `${result.orderNumber}-${result.productType}`;
      if (!seenOrders.has(orderKey)) {
        seenOrders.add(orderKey);
        uniqueResults.push(result);
      }
    }

    console.log(`Final unified results count: ${uniqueResults.length}`);

    if (uniqueResults.length === 0) {
      let responseMessage = isFollowUp ? `*ADDITIONAL ORDER SEARCH*\n\n` : `*ORDER SEARCH RESULTS*\n\n`;
      responseMessage += `No orders found for the given search terms across all product types.\n\n`;
      responseMessage += `Type /menu for main menu or / to go back`;

      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      return;
    }

    // If 5 or fewer results, show in WhatsApp message
    if (uniqueResults.length <= 5) {
      let responseMessage = isFollowUp ? `*ADDITIONAL ORDER SEARCH*\n\n` : `*ORDER SEARCH RESULTS*\n\n`;

      // Group results by product type for better display
      const productGroups = {};
      uniqueResults.forEach(result => {
        if (!productGroups[result.productType]) {
          productGroups[result.productType] = [];
        }
        productGroups[result.productType].push(result);
      });

      Object.entries(productGroups).forEach(([productType, orders]) => {
        responseMessage += `*${productType.toUpperCase()} ORDERS*\n`;
        orders.forEach(result => {
          responseMessage += `*${result.orderNumber}*\n`;
          responseMessage += `${result.message}\n\n`;
        });
      });

      orderQueryTimestamps[from] = Date.now();

      responseMessage += `You can search additional orders within the next 2 minutes by typing order numbers.\n\n`;
      responseMessage += `Type /menu for main menu or / to go back`;

      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);

      userStates[from] = { 
        currentMenu: 'order_followup', 
        timestamp: Date.now()
      };

    } else {
      // More than 5 results - notify user
      await sendWhatsAppMessage(from, `*Large Results Found*

Found ${uniqueResults.length} orders across all product types.
Results are too many for WhatsApp display.

Please use more specific search terms or contact support for detailed results.

Type /menu for main menu`, productId, phoneId);

      delete userStates[from];
    }

  } catch (error) {
    console.error('Error processing unified order query:', error);
    await sendWhatsAppMessage(from, `Error searching orders

Please try again later.

Type /menu for main menu`, productId, phoneId);

    delete userStates[from];
  }
}
// Complete Stock Functions (UNCHANGED)
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
    if (!rows || rows.length === 0) {
      return [];
    }

    const permittedStores = [];

    const phoneVariations = [
      phoneNumber,
      phoneNumber.replace(/^\+91/, ''),
      phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''),
      phoneNumber.replace(/^0/, ''),
      phoneNumber.slice(-10)
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
        if (!rows || rows.length === 0) {
          continue;
        }

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

      termResults.forEach(result => {
        const stockValue = parseFloat(result.stock);
        if (isNaN(stockValue) || stockValue <= 0) {
          return;
        }

        totalResults++;

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
      if (items.length === 0) return;

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

    const message = `*Results Generated*

File: ${filename}
Size: ${fileSizeKB} KB

Download your PDF:
${downloadUrl}
${orderFormMessage}

Click the link above to download
Works on mobile and desktop
Link expires in 5 minutes

Type /menu for main menu`;

    await sendWhatsAppMessage(to, message, productId, phoneId);

  } catch (error) {
    console.error('Error creating download link:', error);
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
      await sendWhatsAppMessage(from, `*Invalid Search*

Please provide at least 5 characters for searching.

Examples:
- 11010 (finds 11010088471-001)
- ABC12 (finds ABC123456789)
- 88471 (finds 11010088471-001)

You can search again within 40 seconds or type /menu for main menu`, productId, phoneId);

      userStates[from] = { 
        currentMenu: 'smart_stock_query',
        lastActivity: Date.now()
      };
      return;
    }

    await sendWhatsAppMessage(from, `*Smart Stock Search*

Searching for: ${validTerms.join(', ')}

Please wait while I search all stock sheets...`, productId, phoneId);

    const searchResults = await searchStockWithPartialMatch(validTerms);
    const permittedStores = await getUserPermittedStores(from);

    let totalResults = 0;
    validTerms.forEach(term => {
      totalResults += (searchResults[term] || []).length;
    });

    if (totalResults === 0) {
      const noResultsMessage = `*No Results Found*

No stock items found containing:
${validTerms.map(term => `- ${term}`).join('\n')}

Try:
- Different search combinations
- Shorter terms (5+ characters)
- Both letters and numbers work

You can search again within 40 seconds or type /menu for main menu`;

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

        const summaryMessage = `*Large Results Found*

Search: ${validTerms.join(', ')}
Total Results: ${totalResults} items
PDF Generated: ${pdfResult.filename}

Results are too long for WhatsApp
PDF does NOT include order forms—see WhatsApp message for order form link

Search more items within 40 seconds or type /menu for main menu`;

        await sendWhatsAppMessage(from, summaryMessage, productId, phoneId);
        await sendWhatsAppFile(from, pdfResult.filepath, pdfResult.filename, productId, phoneId, permittedStores);

        userStates[from] = { 
          currentMenu: 'smart_stock_query',
          lastActivity: Date.now()
        };

      } catch (pdfError) {
        console.error('PDF generation failed:', pdfError);
        await sendWhatsAppMessage(from, `*Error Generating PDF*

Found ${totalResults} results but could not generate PDF.
Please contact support.

Type /menu for main menu`, productId, phoneId);

        delete userStates[from];
      }
    }

  } catch (error) {
    console.error('Error in smart stock query:', error);
    await sendWhatsAppMessage(from, `*Search Error*

Unable to complete search.
Please try again later.

Type /menu for main menu`, productId, phoneId);

    delete userStates[from];
  }
}

// WhatsApp message sending function (UNCHANGED)
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

// MAIN WEBHOOK HANDLER - MODIFIED FOR DIRECT ORDER INPUT
app.post('/webhook', async (req, res) => {
  const message = req.body.message?.text;
  const from = req.body.user?.phone;
  const productId = req.body.product_id || req.body.productId;
  const phoneId = req.body.phone_id || req.body.phoneId;

  if (!message || typeof message !== 'string') {
    return res.sendStatus(200);
  }

  const trimmedMessage = message.trim();
  if (trimmedMessage === '') {
    return res.sendStatus(200);
  }

  // Check for expired stock sessions and clean them up
  if (userStates[from] && isStockSessionExpired(userStates[from])) {
    console.log(`Stock session expired for ${from} at ${new Date().toLocaleTimeString()}`);
    delete userStates[from];
  }

  // Check if this is a valid bot interaction
  if (!isValidBotInteraction(trimmedMessage, userStates[from])) {
    return res.sendStatus(200);
  }

  const lowerMessage = trimmedMessage.toLowerCase();

  // Debug commands (UNCHANGED)
  if (lowerMessage === '/debuggreet') {
    const greeting = await getUserGreeting(from);
    const debugMessage = greeting 
      ? `Found: ${greeting.salutation} ${greeting.name} - ${greeting.greetings}`
      : 'No greeting found in sheet';

    await sendWhatsAppMessage(from, `Greeting debug result: ${debugMessage}`, productId, phoneId);
    return res.sendStatus(200);
  }

  if (lowerMessage === '/debugpermissions') {
    const permissions = await getUserPermissions(from);
    const debugMessage = permissions.length > 0 
      ? `Your permissions: ${permissions.join(', ')}`
      : 'No permissions found';

    await sendWhatsAppMessage(from, `Permission debug result: ${debugMessage}`, productId, phoneId);
    return res.sendStatus(200);
  }

  if (lowerMessage === '/debugtrouser') {
    try {
      const auth = await getGoogleAuth();
      const authClient = await auth.getClient();
      const sheets = google.sheets({ version: 'v4', auth: authClient });

      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: TROUSER_LIVE_SHEET_ID,
        range: `${TROUSER_LIVE_SHEET_NAME}!A1:F10`,
        valueRenderOption: 'UNFORMATTED_VALUE'
      });

      const rows = response.data.values;
      let debugMessage = `*TROUSER SHEET DEBUG*\n\nFirst ${Math.min(10, rows.length)} rows:\n\n`;

      rows.forEach((row, index) => {
        const colD = row.length > 3 ? row[3] : 'EMPTY';
        const colE = row.length > 4 ? row[4] : 'EMPTY';
        debugMessage += `Row ${index}: ColD="${colD}" ColE="${colE}"\n`;
      });

      await sendWhatsAppMessage(from, debugMessage, productId, phoneId);

    } catch (error) {
      await sendWhatsAppMessage(from, `Debug error: ${error.message}`, productId, phoneId);
    }

    return res.sendStatus(200);
  }

  // Handle single "/" for going back one step (MODIFIED)
  if (trimmedMessage === '/') {
    const wentBack = goBackOneStep(from);
    if (wentBack) {
      if (userStates[from] && userStates[from].currentMenu === 'main') {
        const greeting = await getUserGreeting(from);
        const personalizedMenu = await generatePersonalizedMenu(from);
        const finalMessage = formatGreetingMessage(greeting, personalizedMenu);
        await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
      }
      // REMOVED: order_query case since we skip that menu
    } else {
      await sendWhatsAppMessage(from, 'Cannot go back further. Type /menu for main menu.', productId, phoneId);
    }
    return res.sendStatus(200);
  }

  // CONTEXT SWITCHING SHORTCUTS

  // Main menu (UNCHANGED)
  if (lowerMessage === '/menu') {
    userStates[from] = { currentMenu: 'main', timestamp: Date.now() };

    const greeting = await getUserGreeting(from);
    const personalizedMenu = await generatePersonalizedMenu(from);

    const finalMessage = formatGreetingMessage(greeting, personalizedMenu);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Direct Stock Query shortcut (UNCHANGED)
  if (lowerMessage === '/stock') {
    if (!(await hasFeatureAccess(from, 'stock'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Stock Query.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    userStates[from] = { 
      currentMenu: 'smart_stock_query', 
      lastActivity: Date.now()
    };

    const greeting = await getUserGreeting(from);

    const stockQueryPrompt = `*SMART STOCK QUERY*

Enter any 5+ character code (letters/numbers):

Examples:
- 11010 (finds 11010088471-001)
- ABC12 (finds ABC123456-XYZ)  
- 88471 (finds 11010088471-001)

Multiple searches: Separate with commas
Example: 11010, ABC12, 88471

Smart search finds partial matches

Type your search terms below or / to go back:`;

    const finalMessage = formatGreetingMessage(greeting, stockQueryPrompt);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // MODIFIED: Direct Order Query shortcut - No more category selection
  if (lowerMessage === '/order') {
    if (!(await hasFeatureAccess(from, 'order'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Order Query.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    userStates[from] = { currentMenu: 'order_number_input', timestamp: Date.now() };

    const greeting = await getUserGreeting(from);

    const orderQueryPrompt = `*UNIFIED ORDER SEARCH*

Enter your Order Number(s) or search terms:

Full orders: B-J3005Z-1-1, GT54695O-1-1, TR54695O-1-1
Partial search: J3005Z, GT546, TR546, 1234
Multiple: J3005Z, GT546, TR546, ABC123

🔍 Searches ALL product types automatically:
• Shirting orders
• Jacket orders  
• Trouser orders

Type your search terms below or / to go back:`;

    const finalMessage = formatGreetingMessage(greeting, orderQueryPrompt);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // MODIFIED: Order category shortcuts - Now go direct to unified search
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
    userStates[from] = { currentMenu: 'order_number_input', preferredType: category, timestamp: Date.now() };

    const greeting = await getUserGreeting(from);

    const orderQuery = `*UNIFIED ORDER SEARCH*

Enter your Order Number(s) or search terms:

Full orders: ${category === 'Shirting' ? 'B-J3005Z-1-1' : category === 'Jacket' ? 'GT54695O-1-1, D47727S-1-2' : 'TR54695O-1-1'}
Partial search: ${category === 'Shirting' ? 'J3005Z, J300, 1234' : category === 'Jacket' ? 'GT546, D477, 1234' : 'TR546, 1234'}
Multiple: ${category === 'Shirting' ? 'J3005Z, J300, ABC123' : category === 'Jacket' ? 'GT546, D477, ABC123' : 'TR546, ABC123'}

🔍 Searches ALL product types automatically:
• Shirting orders
• Jacket orders
• Trouser orders

Type your search terms below or / to go back:`;

    const finalMessage = formatGreetingMessage(greeting, orderQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Direct ticket shortcuts (UNCHANGED)
  if (lowerMessage === '/helpticket') {
    if (!(await hasFeatureAccess(from, 'help_ticket'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Help Ticket.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    const helpTicketMessage = `*HELP TICKET*

Click the link below to access Help Ticket form directly:

${links.helpTicket}

Type /menu for main menu or / to go back`;

    await sendWhatsAppMessage(from, helpTicketMessage, productId, phoneId);
    delete userStates[from];
    return res.sendStatus(200);
  }

  if (lowerMessage === '/delegation') {
    if (!(await hasFeatureAccess(from, 'delegation'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Delegation.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    const delegationMessage = `*DELEGATION*

Click the link below to access Delegation form directly:

${links.delegation}

Type /menu for main menu or / to go back`;

    await sendWhatsAppMessage(from, delegationMessage, productId, phoneId);
    delete userStates[from];
    return res.sendStatus(200);
  }

  // Handle order follow-up queries within 2-minute window (MODIFIED)
  if (userStates[from] && userStates[from].currentMenu === 'order_followup') {
    if (isWithinOrderQueryWindow(from) && trimmedMessage !== '/menu' && trimmedMessage !== '/') {
      const orderNumbers = trimmedMessage.split(',').map(order => order.trim()).filter(order => order.length > 0);

      if (orderNumbers.length > 0) {
        await processOrderQuery(from, orderNumbers, productId, phoneId, true); // MODIFIED: Removed category
        return res.sendStatus(200);
      }
    } else {
      userStates[from] = { currentMenu: 'main', timestamp: Date.now() };
    }
  }

  // Handle menu selections (MODIFIED - Removed option 2 order category selection)
  if (userStates[from] && userStates[from].currentMenu === 'main') {

    // TICKET OPTION (1) (UNCHANGED)
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
      delete userStates[from];
      return res.sendStatus(200);
    }

    // MODIFIED: ORDER QUERY OPTION (2) - Go direct to unified search
    if (trimmedMessage === '2' && (await hasFeatureAccess(from, 'order'))) {
      userStates[from] = { currentMenu: 'order_number_input', timestamp: Date.now() };

      const greeting = await getUserGreeting(from);

      const orderQueryPrompt = `*UNIFIED ORDER SEARCH*

Enter your Order Number(s) or search terms:

Full orders: B-J3005Z-1-1, GT54695O-1-1, TR54695O-1-1
Partial search: J3005Z, GT546, TR546, 1234
Multiple: J3005Z, GT546, TR546, ABC123

🔍 Searches ALL product types automatically:
• Shirting orders
• Jacket orders
• Trouser orders

Type your search terms below or / to go back:`;

      const finalMessage = formatGreetingMessage(greeting, orderQueryPrompt);
      await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
      return res.sendStatus(200);
    }

    // STOCK QUERY OPTION (3) (UNCHANGED)
    if (trimmedMessage === '3' && (await hasFeatureAccess(from, 'stock'))) {
      userStates[from] = { 
        currentMenu: 'smart_stock_query', 
        lastActivity: Date.now()
      };

      const greeting = await getUserGreeting(from);

      const stockQueryPrompt = `*SMART STOCK QUERY*

Enter any 5+ character code (letters/numbers):

Examples:
- 11010 (finds 11010088471-001)
- ABC12 (finds ABC123456-XYZ)
- 88471 (finds 11010088471-001)

Multiple searches: Separate with commas
Example: 11010, ABC12, 88471

Smart search finds partial matches

Type your search terms below or / to go back:`;

      const finalMessage = formatGreetingMessage(greeting, stockQueryPrompt);
      await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
      return res.sendStatus(200);
    }

    // DOCUMENT OPTION (4) (UNCHANGED)
    if (trimmedMessage === '4' && (await hasFeatureAccess(from, 'document'))) {
      await sendWhatsAppMessage(from, '*DOCUMENT*\n\nThis feature is coming soon!\n\nType /menu for main menu or / to go back', productId, phoneId);
      delete userStates[from];
      return res.sendStatus(200);
    }

    await sendWhatsAppMessage(from, '*ACCESS DENIED*\n\nYou don\'t have permission for this option or invalid selection.\n\nType /menu to see your available options.', productId, phoneId);
    return res.sendStatus(200);
  }

  // REMOVED: Order query category selection handler (no longer needed)

  // MODIFIED: Handle order number input - Now uses unified search
  if (userStates[from] && userStates[from].currentMenu === 'order_number_input') {
    if (trimmedMessage !== '/menu' && trimmedMessage !== '/') {
      const orderNumbers = trimmedMessage.split(',').map(order => order.trim()).filter(order => order.length > 0);

      await processOrderQuery(from, orderNumbers, productId, phoneId); // MODIFIED: Removed category parameter
      return res.sendStatus(200);
    }
  }

  // Handle smart stock query input (UNCHANGED)
  if (userStates[from] && userStates[from].currentMenu === 'smart_stock_query') {
    if (trimmedMessage !== '/menu' && trimmedMessage !== '/') {
      const searchTerms = trimmedMessage.split(',').map(q => q.trim()).filter(q => q.length > 0);
      await processSmartStockQuery(from, searchTerms, productId, phoneId);
      return res.sendStatus(200);
    }
  }

  return res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`WhatsApp Bot running on port ${PORT}`);
  console.log('✅ MAJOR UPDATE: Direct Order Input System Activated');
  console.log('✅ REMOVED: Order category selection menu (1, 2, 3)');
  console.log('✅ NEW: Unified search across ALL product types simultaneously');
  console.log('✅ NEW: Super fast parallel searching with Promise.all()');
  console.log('✅ NEW: Smart result grouping by product type');
  console.log('✅ ENHANCED: /order now goes directly to unified search');
  console.log('✅ All other functions remain completely intact');
  console.log('');
  console.log('🚀 DIRECT ORDER COMMANDS:');
  console.log('   /order - Direct unified order search (from anywhere)');
  console.log('   /shirting - Direct unified search with shirting context');
  console.log('   /jacket - Direct unified search with jacket context'); 
  console.log('   /trouser - Direct unified search with trouser context');
  console.log('');
  console.log('⚡ SPEED IMPROVEMENTS:');
  console.log('   - Parallel searching across Shirting, Jacket & Trouser');
  console.log('   - No menu navigation delays');
  console.log('   - Direct order number input');
  console.log('   - Instant results grouping by product type');
  console.log('');
  console.log('🔍 UNIFIED SEARCH FEATURES:');
  console.log('   - Searches ALL sheets simultaneously');
  console.log('   - Groups results by product type automatically');
  console.log('   - Handles partial and exact matches');
  console.log('   - Removes duplicate results intelligently');
  console.log('   - Displays up to 5 results in WhatsApp');
  console.log('');
  console.log('📱 USER EXPERIENCE:');
  console.log('   - Type /order -> directly enter order numbers');
  console.log('   - No more "1. Shirting 2. Jacket 3. Trouser" menu');
  console.log('   - Get results from all product types instantly');
  console.log('   - Clear product type labels in results');
});

