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
            <p>Please generate a new stock query.</p>
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

// Store order query timestamps for 2-minute window
let orderQueryTimestamps = {};

// Session timeout configuration (40 seconds for stock queries)
const STOCK_SESSION_TIMEOUT = 40 * 1000; // 40 seconds in milliseconds

// Shortcut commands that can switch contexts immediately
const SHORTCUT_COMMANDS = [
  '/menu', '/stock', '/order', '/shirting', '/jacket', '/trouser',
  '/helpticket', '/delegation', '/',
  '/debuggreet', '/debugpermissions', '/debugrows'
];

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
const GREETINGS_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';
const USER_ACCESS_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';

// Static Google Form configuration
const STATIC_FORM_BASE_URL = 'https://docs.google.com/forms/d/e/1FAIpQLSfyAo7LwYtQDfNVxPRbHdk_ymGpDs-RyWTCgzd2PdRhj0T3Hw/viewform';

// Order Query Configuration
const LIVE_SHEET_ID = '1AxjCHsMxYUmEULaW1LxkW78g0Bv9fp4PkZteJO82uEA';
const LIVE_SHEET_NAME = 'FMS';
const COMPLETED_ORDER_FOLDER_ID = '1kgdPdnUK-FsnKZDE5yW6vtRf2H9d3YRE';

// Production stages configuration
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
  { name: 'Dispatch (HO)', column: 'CL', nextStage: 'COMPLETED', dispatchDateColumn: 'CH' }
];

// Function to check if stock session has expired
function isStockSessionExpired(userState) {
  if (!userState || userState.currentMenu !== 'smart_stock_query' || !userState.lastActivity) {
    return false;
  }
  
  const now = Date.now();
  const timeSinceLastActivity = now - userState.lastActivity;
  
  return timeSinceLastActivity > STOCK_SESSION_TIMEOUT;
}

// Function to update last activity timestamp
function updateLastActivity(from) {
  if (userStates[from] && userStates[from].currentMenu === 'smart_stock_query') {
    userStates[from].lastActivity = Date.now();
    console.log(`Updated stock session activity for ${from} at ${new Date().toLocaleTimeString()}`);
  }
}

// UPDATED: Define valid commands and interactions with all shortcuts
function isValidBotInteraction(message, userState) {
  const lowerMessage = message.toLowerCase().trim();
  
  // Check for debug commands with parameters
  if (lowerMessage.startsWith('/debugorder ')) {
    return true;
  }
  
  // Check for shortcut commands - these should ALWAYS work regardless of state
  if (SHORTCUT_COMMANDS.includes(lowerMessage)) {
    return true;
  }
  
  // Check for numbered menu selections (1, 2, 3, 4)
  if (['1', '2', '3', '4'].includes(message.trim())) {
    return true;
  }
  
  // Check user state and session expiry
  if (userState) {
    switch (userState.currentMenu) {
      case 'main':
        return ['1', '2', '3', '4'].includes(message.trim());
        
      case 'order_query':
        return ['1', '2', '3'].includes(message.trim());
        
      case 'order_number_input':
      case 'order_followup':
        return message.trim().length > 0;
        
      case 'smart_stock_query':
        // Check if stock session has expired
        if (isStockSessionExpired(userState)) {
          return false; // Session expired, ignore the message
        }
        return message.trim().length > 0;
        
      default:
        return false;
    }
  }
  
  // If no active state, ignore casual messages
  return false;
}

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

// Format date for display
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

// Format stock quantity - show 15+ if > 15
function formatStockQuantity(stockValue) {
  if (!stockValue || stockValue === '') return stockValue;
  
  const numValue = parseFloat(stockValue.toString().trim());
  if (!isNaN(numValue) && numValue > 15) {
    return '15+';
  }
  
  return stockValue.toString();
}

// Helper function to go back one step
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
  
  if (currentMenu === 'main' || currentMenu === 'completed') {
    return false;
  }
  
  return false;
}

// Get user permissions from BOT Permission sheet
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

// Check if user has access to specific feature
async function hasFeatureAccess(phoneNumber, feature) {
  const userPermissions = await getUserPermissions(phoneNumber);
  return userPermissions.includes(feature.toLowerCase());
}

// Generate personalized menu with new shortcuts
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
    shortcuts.push('/order - Direct Order Menu');
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

// Generate personalized ticket menu based on specific permissions
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

// Get user greeting from separate columns
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

// Format greeting message
function formatGreetingMessage(greeting, mainMessage) {
  if (!greeting || !greeting.name || !greeting.salutation || !greeting.greetings) {
    return mainMessage;
  }
  
  return `${greeting.salutation} ${greeting.name}\n\n${greeting.greetings}\n\n${mainMessage}`;
}

// Handle separate columns for stock data with duplicate removal and max quantity selection
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

// Generate PDF with proper format
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

// NEW: Generate Order Results PDF
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

// Send file via Railway
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

// Get user permitted stores from separate columns
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

// UPDATED: Smart Stock Query with 40-second session management
async function processSmartStockQuery(from, searchTerms, productId, phoneId) {
  try {
    // Update activity timestamp for this user
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
      
      // Keep user in stock query state but update activity
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
      
      // Keep user in stock query state for 40 more seconds
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
      
      // Keep user in stock query state for 40 more seconds
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
        
        // Keep user in stock query state for 40 more seconds
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
        
        // Reset state on error
        delete userStates[from];
      }
    }
    
  } catch (error) {
    console.error('Error in smart stock query:', error);
    await sendWhatsAppMessage(from, `*Search Error*

Unable to complete search.
Please try again later.

Type /menu for main menu`, productId, phoneId);
    
    // Reset state on error
    delete userStates[from];
  }
}

// Convert column letter to index (A=0, B=1, etc.)
function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

// Check production stages with date formatting
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

    if (!hasAnyStage) {
      return { message: 'Order is currently under process' };
    }

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
      return { 
        message: `Order is currently completed ${lastCompletedStage.name} stage and processed to ${lastCompletedStage.nextStage} stage` 
      };
    }

    return { message: 'Error determining order status' };

  } catch (error) {
    console.error('Error checking production stages:', error);
    return { message: 'Error checking order status' };
  }
}

// NEW: Search orders by partial match (for 4-digit search)
async function searchOrdersByPartialMatch(searchTerm) {
  const matchingOrders = [];
  
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

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
        
        if (orderNumber && orderNumber.toUpperCase().includes(searchTerm.toUpperCase())) {
          const stageStatus = checkProductionStages(row);
          matchingOrders.push({
            orderNumber: orderNumber,
            message: stageStatus.message,
            location: 'Live Sheet (FMS)'
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
          
          if (orderNumber && orderNumber.toUpperCase().includes(searchTerm.toUpperCase())) {
            let rawDispatchDate = '';
            if (row.length > 90 && row[90] !== undefined && row[90] !== null) {
              rawDispatchDate = row[90];
            }
            
            const formattedDate = formatDateForDisplay(rawDispatchDate);
            
            matchingOrders.push({
              orderNumber: orderNumber,
              message: `Order got dispatched on ${formattedDate}`,
              location: 'Completed Orders'
            });
          }
        }
      } catch (error) {
        console.error(`Error searching ${file.name}:`, error.message);
        continue;
      }
    }

    return matchingOrders;

  } catch (error) {
    console.error('Error in searchOrdersByPartialMatch:', error);
    return [];
  }
}

// Search in live sheet with enhanced debugging
async function searchInLiveSheet(sheets, orderNumber) {
  try {
    console.log(`=== SEARCHING FOR ORDER: ${orderNumber} ===`);
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: LIVE_SHEET_ID,
      range: `${LIVE_SHEET_NAME}!A:CL`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log('ERROR: No rows returned from live sheet');
      return { found: false };
    }

    console.log(`Live sheet has ${rows.length} total rows`);

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      
      let sheetOrderNumber = '';
      if (row.length > 3 && row[3] !== undefined && row[3] !== null) {
        sheetOrderNumber = row[3].toString().trim();
      }
      
      if (sheetOrderNumber) {
        console.log(`Row ${i}: Found order "${sheetOrderNumber}" (length: ${sheetOrderNumber.length})`);
        
        const searchOrder = orderNumber.trim().toUpperCase();
        const sheetOrder = sheetOrderNumber.toUpperCase();
        
        console.log(`Comparing: "${searchOrder}" === "${sheetOrder}"`);
        
        if (sheetOrder === searchOrder) {
          console.log(`✅ EXACT MATCH FOUND at row ${i}`);
          const stageStatus = checkProductionStages(row);
          return {
            found: true,
            message: stageStatus.message,
            location: 'Live Sheet (FMS)'
          };
        }
      }
    }

    console.log(`❌ Order ${orderNumber} not found in any of the ${rows.length} rows`);
    return { found: false };

  } catch (error) {
    console.error('Error searching live sheet:', error.message);
    return { found: false };
  }
}

// Search in completed orders with date formatting
async function searchInCompletedSheetSimplified(sheets, sheetId, orderNumber) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'A:CL',
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return { found: false };
    }

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

// Main order search function
async function searchOrderStatus(orderNumber, category) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    const liveSheetResult = await searchInLiveSheet(sheets, orderNumber);
    if (liveSheetResult.found) {
      return liveSheetResult;
    }

    const folderFiles = await drive.files.list({
      q: `'${COMPLETED_ORDER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    for (const file of folderFiles.data.files) {
      try {
        const completedResult = await searchInCompletedSheetSimplified(sheets, file.id, orderNumber);
        if (completedResult.found) {
          return completedResult;
        }
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

// Check if user is within 2-minute order query window
function isWithinOrderQueryWindow(from) {
  if (!orderQueryTimestamps[from]) return false;
  
  const now = Date.now();
  const lastQuery = orderQueryTimestamps[from];
  const twoMinutes = 2 * 60 * 1000;
  
  return (now - lastQuery) < twoMinutes;
}

// UPDATED: Process order query with 4-digit search and PDF support
async function processOrderQuery(from, category, orderNumbers, productId, phoneId, isFollowUp = false) {
  try {
    if (!isFollowUp) {
      await sendWhatsAppMessage(from, `*Checking ${category} orders...*

Please wait while I search for your order status.`, productId, phoneId);
    }

    let orderResults = [];
    let exactMatches = [];
    let partialMatches = [];

    for (const orderInput of orderNumbers) {
      const cleanInput = orderInput.trim();
      
      // Check if it's a 4-digit search
      if (cleanInput.length === 4 && /^\d{4}$/.test(cleanInput)) {
        // 4-digit search - find all matching orders
        const matchingOrders = await searchOrdersByPartialMatch(cleanInput);
        partialMatches = partialMatches.concat(matchingOrders);
      } else {
        // Exact order number search
        const orderStatus = await searchOrderStatus(cleanInput, category);
        exactMatches.push({
          orderNumber: cleanInput,
          message: orderStatus.message,
          location: orderStatus.location || 'Unknown'
        });
      }
    }

    // Combine all results
    orderResults = [...exactMatches, ...partialMatches];

    if (orderResults.length === 0) {
      let responseMessage = isFollowUp ? `*ADDITIONAL ${category.toUpperCase()} ORDER STATUS*\n\n` : `*${category.toUpperCase()} ORDER STATUS*\n\n`;
      responseMessage += `No orders found for the given search terms.\n\n`;
      responseMessage += `Type /menu for main menu or / to go back`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      return;
    }

    // If 3 or fewer results, show in WhatsApp message
    if (orderResults.length <= 3) {
      let responseMessage = isFollowUp ? `*ADDITIONAL ${category.toUpperCase()} ORDER STATUS*\n\n` : `*${category.toUpperCase()} ORDER STATUS*\n\n`;
      
      orderResults.forEach(result => {
        responseMessage += `*Order: ${result.orderNumber}*\n`;
        responseMessage += `${result.message}\n\n`;
      });
      
      orderQueryTimestamps[from] = Date.now();
      
      responseMessage += `You can query additional ${category} orders within the next 2 minutes by simply typing the order numbers.\n\n`;
      responseMessage += `Type /menu for main menu or / to go back`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      
      userStates[from] = { 
        currentMenu: 'order_followup', 
        category: category,
        timestamp: Date.now()
      };
      
    } else {
      // More than 3 results - generate PDF
      try {
        const searchTerm = orderNumbers.join(', ');
        const pdfResult = await generateOrderPDF(orderResults, searchTerm, from, category);
        
        const summaryMessage = `*Large Results Found*

Search: ${searchTerm}
Total Results: ${orderResults.length} orders
PDF Generated: ${pdfResult.filename}

Results are too many for WhatsApp

Type /menu for main menu`;
        
        await sendWhatsAppMessage(from, summaryMessage, productId, phoneId);
        await sendWhatsAppFile(from, pdfResult.filepath, pdfResult.filename, productId, phoneId, []);
        
        // Reset state after PDF generation
        delete userStates[from];
        
      } catch (pdfError) {
        console.error('PDF generation failed:', pdfError);
        await sendWhatsAppMessage(from, `*Error Generating PDF*

Found ${orderResults.length} results but could not generate PDF.
Please contact support.

Type /menu for main menu`, productId, phoneId);
        
        // Reset state on error
        delete userStates[from];
      }
    }
    
  } catch (error) {
    console.error('Error processing order query:', error);
    await sendWhatsAppMessage(from, 'Error checking orders\n\nPlease try again later.\n\nType /menu for main menu', productId, phoneId);
    
    // Reset state on error
    delete userStates[from];
  }
}

// MAIN WEBHOOK HANDLER - WITH ALL CONTEXT SWITCHING SHORTCUTS
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
    // IGNORE: Don't respond to random messages
    return res.sendStatus(200);
  }

  const lowerMessage = trimmedMessage.toLowerCase();

  // Debug commands
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

  if (lowerMessage.startsWith('/debugorder ')) {
    const testOrderNumber = trimmedMessage.replace('/debugorder ', '').trim();
    
    await sendWhatsAppMessage(from, `🔍 Debugging order search for: ${testOrderNumber}`, productId, phoneId);
    
    try {
      const auth = await getGoogleAuth();
      const authClient = await auth.getClient();
      const sheets = google.sheets({ version: 'v4', auth: authClient });
      
      const result = await searchInLiveSheet(sheets, testOrderNumber);
      
      const debugResult = result.found 
        ? `✅ Order FOUND: ${result.message}`
        : `❌ Order NOT FOUND in live sheet`;
        
      await sendWhatsAppMessage(from, debugResult, productId, phoneId);
      
    } catch (error) {
      await sendWhatsAppMessage(from, `Error during debug: ${error.message}`, productId, phoneId);
    }
    
    return res.sendStatus(200);
  }

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
        const orderQueryMenu = `*ORDER QUERY*

Please select the product category:

1. Shirting
2. Jacket  
3. Trouser

Type the number to continue or / to go back`;
        
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

  // NEW: Direct Order Query shortcut
  if (lowerMessage === '/order') {
    if (!(await hasFeatureAccess(from, 'order'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Order Query.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    userStates[from] = { currentMenu: 'order_query', timestamp: Date.now() };
    
    const greeting = await getUserGreeting(from);
    
    const orderQueryMenu = `*ORDER QUERY*

Please select the product category:

1. Shirting
2. Jacket  
3. Trouser

Type the number to continue or / to go back`;
    
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
    
    const orderQuery = `*${category.toUpperCase()} ORDER QUERY*

Please enter your Order Number(s) or 4-digit search:

Full order: ABC123DEF456
4-digit search: 1234 (finds all orders containing 1234)
Multiple: ABC123, 1234, DEF456

Type your search terms below or / to go back:`;
    
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

    const helpTicketMessage = `*HELP TICKET*

Click the link below to access Help Ticket form directly:

${links.helpTicket}

Type /menu for main menu or / to go back`;

    await sendWhatsAppMessage(from, helpTicketMessage, productId, phoneId);
    delete userStates[from]; // Clear any existing state
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
      const orderQueryMenu = `*ORDER QUERY*

Please select the product category:

1. Shirting
2. Jacket  
3. Trouser

Type the number to continue or / to go back`;
      
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
      await sendWhatsAppMessage(from, `*SHIRTING ORDER QUERY*

Please enter your Order Number(s) or 4-digit search:

Full order: ABC123DEF456
4-digit search: 1234 (finds all orders containing 1234)
Multiple: ABC123, 1234, DEF456

Type your search terms below or / to go back:`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Jacket', timestamp: Date.now() };
      await sendWhatsAppMessage(from, `*JACKET ORDER QUERY*

Please enter your Order Number(s) or 4-digit search:

Full order: ABC123DEF456
4-digit search: 1234 (finds all orders containing 1234)
Multiple: ABC123, 1234, DEF456

Type your search terms below or / to go back:`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Trouser', timestamp: Date.now() };
      await sendWhatsAppMessage(from, `*TROUSER ORDER QUERY*

Please enter your Order Number(s) or 4-digit search:

Full order: ABC123DEF456
4-digit search: 1234 (finds all orders containing 1234)
Multiple: ABC123, 1234, DEF456

Type your search terms below or / to go back:`, productId, phoneId);
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
  console.log('✅ NEW: 4-digit order search with PDF generation for >3 results');
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
  console.log('🔍 ORDER SEARCH FEATURES:');
  console.log('   - Full order numbers: ABC123DEF456');
  console.log('   - 4-digit search: 1234 (finds all matching orders)');
  console.log('   - PDF generation for >3 results');
  console.log('   - WhatsApp display for ≤3 results');
  console.log('');
  console.log('Debug commands: /debuggreet, /debugpermissions, /debugorder, /debugrows');
});
