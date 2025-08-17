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
// UPDATED: Using your BOT Permission sheet
const USER_ACCESS_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';

// Static Google Form configuration
const STATIC_FORM_BASE_URL = 'https://docs.google.com/forms/d/e/1FAIpQLSfyAo7LwYtQDfNVxPRbHdk_ymGpDs-RyWTCgzd2PdRhj0T3Hw/viewform';

// Order Query Configuration
const LIVE_SHEET_ID = '1AxjCHsMxYUmEULaW1LxkW78g0Bv9fp4PkZteJO82uEA';
const LIVE_SHEET_NAME = 'FMS';
const COMPLETED_ORDER_FOLDER_ID = '1kgdPdnUK-FsnKZDE5yW6vtRf2H9d3YRE';

// Production stages configuration
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

// UPDATED: Get user permissions from BOT Permission sheet
async function getUserPermissions(phoneNumber) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    // Using your sheet name "BOT Permission"
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

      const sheetPhone = (row[0] || '').toString().trim();  // Column A: Contact Number
      const featuresString = (row[1] || '').toString().trim(); // Column B: Features

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

// Generate personalized menu with fine-grained ticket options
async function generatePersonalizedMenu(phoneNumber) {
  const userPermissions = await getUserPermissions(phoneNumber);
  
  if (userPermissions.length === 0) {
    return `*ACCESS DENIED*

You do not have permission to use this bot.
Please contact administrator for access.`;
  }

  let menuItems = [];
  let shortcuts = [];

  // Check for any ticket-related permissions
  const hasAnyTicketAccess = userPermissions.some(perm => 
    ['help_ticket', 'delegation', 'leave_form'].includes(perm)
  );

  if (hasAnyTicketAccess) {
    menuItems.push('1. Ticket');
  }
  
  if (userPermissions.includes('order')) {
    menuItems.push('2. Order Query');
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

  menu += `\n\nType the number or use shortcuts`;

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

Type /menu to return to main menu`;
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

// Handle separate columns for stock data
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

    // Header
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
          stock: result.stock
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

    const message = `*Stock Results Generated*

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

// Smart Stock Query
async function processSmartStockQuery(from, searchTerms, productId, phoneId) {
  try {
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

Type /menu for main menu`, productId, phoneId);
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

Type /menu for main menu`;
      
      await sendWhatsAppMessage(from, noResultsMessage, productId, phoneId);
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
          responseMessage += `${item.qualityCode}: ${item.stock}\n`;
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
      
      responseMessage += `Type /menu for main menu`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      
    } else {
      try {
        const pdfResult = await generateStockPDF(searchResults, validTerms, from, permittedStores);
        
        const summaryMessage = `*Large Results Found*

Search: ${validTerms.join(', ')}
Total Results: ${totalResults} items
PDF Generated: ${pdfResult.filename}

Results are too long for WhatsApp
PDF does NOT include order forms—see WhatsApp message for order form link`;
        
        await sendWhatsAppMessage(from, summaryMessage, productId, phoneId);
        await sendWhatsAppFile(from, pdfResult.filepath, pdfResult.filename, productId, phoneId, permittedStores);
        
      } catch (pdfError) {
        console.error('PDF generation failed:', pdfError);
        await sendWhatsAppMessage(from, `*Error Generating PDF*

Found ${totalResults} results but could not generate PDF.
Please contact support.

Type /menu for main menu`, productId, phoneId);
      }
    }
    
  } catch (error) {
    console.error('Error in smart stock query:', error);
    await sendWhatsAppMessage(from, `*Search Error*

Unable to complete search.
Please try again later.

Type /menu for main menu`, productId, phoneId);
  }
}

// Main webhook handler with security checks
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

  const lowerMessage = trimmedMessage.toLowerCase();

  // Debug greeting command
  if (lowerMessage === '/debuggreet') {
    const greeting = await getUserGreeting(from);
    const debugMessage = greeting 
      ? `Found: ${greeting.salutation} ${greeting.name} - ${greeting.greetings}`
      : 'No greeting found in sheet';
    
    await sendWhatsAppMessage(from, `Greeting debug result: ${debugMessage}`, productId, phoneId);
    return res.sendStatus(200);
  }

  // Debug permissions command
  if (lowerMessage === '/debugpermissions') {
    const permissions = await getUserPermissions(from);
    const debugMessage = permissions.length > 0 
      ? `Your permissions: ${permissions.join(', ')}`
      : 'No permissions found';
    
    await sendWhatsAppMessage(from, `Permission debug result: ${debugMessage}`, productId, phoneId);
    return res.sendStatus(200);
  }

  // Main menu with security check
  if (lowerMessage === '/menu' || trimmedMessage === '/') {
    userStates[from] = { currentMenu: 'main' };
    
    const greeting = await getUserGreeting(from);
    const personalizedMenu = await generatePersonalizedMenu(from);
    
    const finalMessage = formatGreetingMessage(greeting, personalizedMenu);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Direct Stock Query shortcut with security check
  if (lowerMessage === '/stock') {
    if (!(await hasFeatureAccess(from, 'stock'))) {
      await sendWhatsAppMessage(from, `*ACCESS DENIED*\n\nYou don't have permission to access Stock Query.\nContact administrator for access.`, productId, phoneId);
      return res.sendStatus(200);
    }

    userStates[from] = { currentMenu: 'smart_stock_query' };
    
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

Type your search terms below:`;
    
    const finalMessage = formatGreetingMessage(greeting, stockQueryPrompt);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Order shortcuts with security check
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
    userStates[from] = { currentMenu: 'order_number_input', category: category };
    
    const greeting = await getUserGreeting(from);
    
    const orderQuery = `*${category.toUpperCase()} ORDER QUERY*

Please enter your Order Number(s):

Single order: ABC123
Multiple orders: ABC123, DEF456, GHI789

Type your order numbers below:`;
    
    const finalMessage = formatGreetingMessage(greeting, orderQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle menu selections with fine-grained security checks
  if (userStates[from] && userStates[from].currentMenu === 'main') {
    
    // TICKET OPTION (1) - Check for any ticket permission
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
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }

    // ORDER QUERY OPTION (2)
    if (trimmedMessage === '2' && (await hasFeatureAccess(from, 'order'))) {
      userStates[from].currentMenu = 'order_query';
      const orderQueryMenu = `*ORDER QUERY*

Please select the product category:

1. Shirting
2. Jacket  
3. Trouser

Type the number to continue`;
      
      await sendWhatsAppMessage(from, orderQueryMenu, productId, phoneId);
      return res.sendStatus(200);
    }

    // STOCK QUERY OPTION (3)
    if (trimmedMessage === '3' && (await hasFeatureAccess(from, 'stock'))) {
      userStates[from] = { currentMenu: 'smart_stock_query' };
      
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

Type your search terms below:`;
      
      const finalMessage = formatGreetingMessage(greeting, stockQueryPrompt);
      await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
      return res.sendStatus(200);
    }

    // DOCUMENT OPTION (4)
    if (trimmedMessage === '4' && (await hasFeatureAccess(from, 'document'))) {
      await sendWhatsAppMessage(from, '*DOCUMENT*\n\nThis feature is coming soon!\n\nType /menu to return to main menu', productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }

    // If user tries to access unauthorized feature
    await sendWhatsAppMessage(from, '*ACCESS DENIED*\n\nYou don\'t have permission for this option or invalid selection.\n\nType /menu to see your available options.', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle order query category selection
  if (userStates[from] && userStates[from].currentMenu === 'order_query') {
    if (trimmedMessage === '1') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Shirting' };
      await sendWhatsAppMessage(from, `*SHIRTING ORDER QUERY*\n\nPlease enter your Order Number(s):\n\nSingle order: ABC123\nMultiple orders: ABC123, DEF456, GHI789\n\nType your order numbers below:`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Jacket' };
      await sendWhatsAppMessage(from, `*JACKET ORDER QUERY*\n\nPlease enter your Order Number(s):\n\nSingle order: ABC123\nMultiple orders: ABC123, DEF456, GHI789\n\nType your order numbers below:`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Trouser' };
      await sendWhatsAppMessage(from, `*TROUSER ORDER QUERY*\n\nPlease enter your Order Number(s):\n\nSingle order: ABC123\nMultiple orders: ABC123, DEF456, GHI789\n\nType your order numbers below:`, productId, phoneId);
      return res.sendStatus(200);
    }

    await sendWhatsAppMessage(from, 'Invalid option. Please select 1, 2, or 3.\n\nType /menu to return to main menu', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle order number input
  if (userStates[from] && userStates[from].currentMenu === 'order_number_input') {
    if (trimmedMessage !== '/menu') {
      const category = userStates[from].category;
      const orderNumbers = trimmedMessage.split(',').map(order => order.trim()).filter(order => order.length > 0);
      
      await processOrderQuery(from, category, orderNumbers, productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }
  }

  // Handle smart stock query input
  if (userStates[from] && userStates[from].currentMenu === 'smart_stock_query') {
    if (trimmedMessage !== '/menu') {
      const searchTerms = trimmedMessage.split(',').map(q => q.trim()).filter(q => q.length > 0);
      await processSmartStockQuery(from, searchTerms, productId, phoneId);
      return res.sendStatus(200);
    }
  }

  return res.sendStatus(200);
});

// Convert column letter to index (A=0, B=1, etc.)
function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

// Check production stages
function checkProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let hasAnyStage = false;

    for (let i = 0; i < PRODUCTION_STAGES.length; i++) {
      const stage = PRODUCTION_STAGES[i];
      const columnIndex = columnToIndex(stage.column);
      const cellValue = row[columnIndex] ? row[columnIndex].toString().trim() : '';
      
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
      const dispatchDate = row[dispatchDateIndex] ? row[dispatchDateIndex].toString().trim() : 'Date not available';
      return { message: `Order has been dispatched from HO on ${dispatchDate}` };
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

// Search in live sheet (FMS)
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

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const sheetOrderNumber = row[3] ? row[3].toString().trim() : '';
      if (!sheetOrderNumber) continue;

      if (sheetOrderNumber.toUpperCase() === orderNumber.trim().toUpperCase()) {
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

// Search in completed orders
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

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const sheetOrderNumber = row[3] ? row[3].toString().trim() : '';
      if (!sheetOrderNumber) continue;

      if (sheetOrderNumber.toUpperCase() === orderNumber.trim().toUpperCase()) {
        const dispatchDate = row[86] ? row[86].toString().trim() : 'Date not available';
        return {
          found: true,
          message: `Order got dispatched on ${dispatchDate}`,
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

// Process order query
async function processOrderQuery(from, category, orderNumbers, productId, phoneId) {
  try {
    await sendWhatsAppMessage(from, `*Checking ${category} orders...*

Please wait while I search for your order status.`, productId, phoneId);

    let responseMessage = `*${category.toUpperCase()} ORDER STATUS*\n\n`;
    
    for (const orderNum of orderNumbers) {
      const orderStatus = await searchOrderStatus(orderNum, category);
      
      responseMessage += `*Order: ${orderNum}*\n`;
      responseMessage += `${orderStatus.message}\n\n`;
    }
    
    responseMessage += `Type /menu to return to main menu`;
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
  } catch (error) {
    console.error('Error processing order query:', error);
    await sendWhatsAppMessage(from, 'Error checking orders\n\nPlease try again later.\n\nType /menu to return to main menu', productId, phoneId);
  }
}

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
  console.log('✅ Fine-Grained Access Control implemented');
  console.log('✅ Using BOT Permission sheet for user access control');
  console.log('✅ Sub-features: help_ticket, delegation, leave_form, order, stock, document');
  console.log('✅ Users see only their permitted options');
  console.log('Debug commands: /debuggreet, /debugpermissions');
  console.log('Sheet: BOT Permission with columns A (Contact Number) | B (Features)');
});
