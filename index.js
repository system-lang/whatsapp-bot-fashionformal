require('dotenv').config();
const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
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
const GREETINGS_SHEET_ID = '1fK1JjsKgdt0tqawUKKgvcrgekj28uvqibk3QIFjtzbE';

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

// Get user greeting from Greetings sheet
async function getUserGreeting(phoneNumber) {
  try {
    console.log(`Getting greeting for phone: ${phoneNumber}`);
    
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    // Get sheet metadata first
    let rows = null;
    
    try {
      const metaResponse = await sheets.spreadsheets.get({
        spreadsheetId: GREETINGS_SHEET_ID,
      });
      
      console.log('Available sheets in spreadsheet:');
      metaResponse.data.sheets.forEach(sheet => {
        console.log(`- Sheet: "${sheet.properties.title}" (ID: ${sheet.properties.sheetId})`);
      });
      
      // Find the greetings sheet (GID 904469862)
      let greetingsSheet = metaResponse.data.sheets.find(sheet => 
        sheet.properties.sheetId === 904469862
      );
      
      if (!greetingsSheet) {
        greetingsSheet = metaResponse.data.sheets.find(sheet => 
          sheet.properties.title.toLowerCase().includes('greet')
        );
      }
      
      if (greetingsSheet) {
        const sheetName = greetingsSheet.properties.title;
        console.log(`Found greetings sheet: "${sheetName}"`);
        
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: GREETINGS_SHEET_ID,
          range: `'${sheetName}'!A:D`,
          valueRenderOption: 'UNFORMATTED_VALUE'
        });
        
        rows = response.data.values;
        console.log(`Successfully read ${rows ? rows.length : 0} rows from greetings sheet`);
      }
    } catch (metaError) {
      console.log('Metadata approach failed, trying alternative ranges...');
      
      // Fallback attempts
      const attempts = [
        'Greetings!A:D',
        "'Greetings'!A:D", 
        'Sheet2!A:D',
        'A:D'
      ];
      
      for (const range of attempts) {
        try {
          console.log(`Trying range: ${range}`);
          const response = await sheets.spreadsheets.values.get({
            spreadsheetId: GREETINGS_SHEET_ID,
            range: range,
            valueRenderOption: 'UNFORMATTED_VALUE'
          });
          
          if (response.data.values && response.data.values.length > 0) {
            rows = response.data.values;
            console.log(`Successfully accessed greetings with range: ${range}`);
            break;
          }
        } catch (rangeError) {
          console.log(`Range ${range} failed: ${rangeError.message}`);
          continue;
        }
      }
    }
    
    console.log('Greetings sheet raw data:', JSON.stringify(rows, null, 2));
    
    if (!rows || rows.length === 0) {
      console.log('No data found in any Greetings sheet attempt');
      return null;
    }

    // Clean phone number variations (remove duplicates and fix length issue)
    const phoneVariations = Array.from(new Set([
      phoneNumber,
      phoneNumber.replace(/^\+91/, ''),
      phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''),
      phoneNumber.replace(/^0/, ''),
      phoneNumber.replace(/[\s\-\(\)]/g, ''),
      phoneNumber.slice(-10) // Last 10 digits
    ])).filter(p => p && p.length >= 10);
    
    console.log(`Looking for greeting with phone variations: ${phoneVariations.join(', ')}`);
    
    // Handle JavaScript numbers and field mapping correctly
    for (let i = 1; i < rows.length; i++) { // Skip header row
      const row = rows[i];
      if (!row || row.length === 0) {
        console.log(`Row ${i + 1}: Skipping empty greeting row`);
        continue;
      }
      
      // Safely convert contact to string (handle JavaScript numbers)
      let sheetContact;
      try {
        const contactCell = row[0];
        if (typeof contactCell === 'number') {
          // Handle JavaScript numbers - convert to string without scientific notation
          sheetContact = contactCell.toString();
          // If it's still in scientific notation, use format
          if (sheetContact.includes('e+') || sheetContact.includes('E+')) {
            sheetContact = contactCell.toFixed(0);
          }
        } else if (contactCell !== null && contactCell !== undefined) {
          sheetContact = contactCell.toString();
          // Handle comma-separated corruption
          if (sheetContact.includes(',')) {
            sheetContact = sheetContact.split(',')[0].trim();
          }
        } else {
          sheetContact = '';
        }
        sheetContact = sheetContact.trim();
      } catch (error) {
        console.log(`Row ${i + 1}: Error processing contact field: ${error.message}`);
        continue;
      }
      
      // Proper field mapping based on your sheet structure
      // Your sheet: Contact Number | Name | Salutation | Greetings
      const name = row[1] ? row[1].toString().trim() : '';
      const salutation = row[1] ? row[1].toString().trim() : '';
      const greetings = row ? row.toString().trim() : '';
      
      console.log(`Row ${i + 1}: FINAL - Contact="${sheetContact}", Name="${name}", Salutation="${salutation}", Greetings="${greetings}"`);
      
      // Clean sheet contact variations (remove duplicates)
      const sheetContactVariations = Array.from(new Set([
        sheetContact,
        sheetContact.replace(/^\+91/, ''),
        sheetContact.replace(/^\+/, ''),
        sheetContact.replace(/^91/, ''),
        sheetContact.replace(/^0/, ''),
        sheetContact.replace(/[\s\-\(\)]/g, ''),
        sheetContact.slice(-10) // Last 10 digits
      ])).filter(s => s && s.length >= 10);
      
      console.log(`Comparing phone variations: ${phoneVariations.join(', ')} with sheet variations: ${sheetContactVariations.join(', ')}`);
      
      // Check for match
      let isMatch = false;
      for (const phoneVar of phoneVariations) {
        if (sheetContactVariations.includes(phoneVar)) {
          console.log(`✅ Greeting match found! "${phoneVar}"`);
          isMatch = true;
          break;
        }
      }
      
      if (isMatch) {
        console.log(`🎉 Found greeting for ${phoneNumber}: ${salutation} ${name}`);
        return {
          name: name,
          salutation: salutation,
          greetings: greetings
        };
      }
    }
    
    console.log(`❌ No greeting found for ${phoneNumber} after checking all rows`);
    return null;
    
  } catch (error) {
    console.error('❌ Error getting user greeting:', error);
    return null;
  }
}

// Format greeting message
function formatGreetingMessage(greeting, mainMessage) {
  if (!greeting) {
    return mainMessage;
  }
  
  return `${greeting.salutation} ${greeting.name}\n\n${greeting.greetings}\n\n${mainMessage}`;
}

// FIXED: Enhanced Smart Stock Search with Alphanumeric Support and Clean Output
async function searchStockWithPartialMatch(searchTerms) {
  const results = {};
  
  // Initialize results for each search term
  searchTerms.forEach(term => {
    results[term] = [];
  });

  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    // Get all stock files
    const folderFiles = await drive.files.list({
      q: `'${STOCK_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
      fields: 'files(id, name)'
    });

    console.log(`Found ${folderFiles.data.files.length} stock files for smart search`);

    // Search through each stock file
    for (const file of folderFiles.data.files) {
      console.log(`Smart searching in: ${file.name}`);
      
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: file.id,
          range: 'A:E', // Only need columns A (code) and E (stock)
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
          console.log(`No data in ${file.name}`);
          continue;
        }
        
        // Search each row for partial matches
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row[0]) continue;
          
          const qualityCode = row[0].toString().trim();
          const stockValue = row ? row.toString().trim() : '0'; // Column E (index 4)
          
          // Check each search term for partial matches
          searchTerms.forEach(searchTerm => {
            const cleanSearchTerm = searchTerm.trim();
            
            // FIXED: Accept any 5+ character alphanumeric substring (not just digits)
            if (cleanSearchTerm.length >= 5 && qualityCode.toUpperCase().includes(cleanSearchTerm.toUpperCase())) {
              console.log(`📍 ALPHANUMERIC MATCH: "${cleanSearchTerm}" found in "${qualityCode}" at ${file.name} - Stock: ${stockValue}`);
              
              results[searchTerm].push({
                qualityCode: qualityCode,
                stock: stockValue,
                store: file.name,
                searchTerm: cleanSearchTerm
              });
            }
          });
        }

      } catch (sheetError) {
        console.error(`Error accessing stock sheet ${file.name}:`, sheetError.message);
      }
    }

    return results;

  } catch (error) {
    console.error('Error in smart stock search:', error);
    throw error;
  }
}

// FIXED: Generate PDF showing only Quality Code and Stock (Column E)
async function generateStockPDF(searchResults, searchTerms, phoneNumber) {
  try {
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    const filename = `stock_results_${phoneNumber.slice(-4)}_${timestamp}.pdf`;
    const filepath = path.join(__dirname, 'temp', filename);
    
    // Ensure temp directory exists
    const tempDir = path.join(__dirname, 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    const doc = new PDFDocument({
      margin: 50,
      size: 'A4'
    });

    // Create write stream
    const stream = fs.createWriteStream(filepath);
    doc.pipe(stream);

    // Header
    doc.fontSize(20)
       .font('Helvetica-Bold')
       .text('Stock Query Results', { align: 'center' });
    
    doc.moveDown(0.5);
    
    doc.fontSize(12)
       .font('Helvetica')
       .text(`Generated: ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}`)
       .text(`Search Terms: ${searchTerms.join(', ')}`)
       .text(`Phone: ${phoneNumber}`)
       .moveDown();

    // Results
    let totalResults = 0;
    
    searchTerms.forEach(term => {
      const termResults = searchResults[term] || [];
      totalResults += termResults.length;
      
      if (termResults.length > 0) {
        doc.fontSize(16)
           .font('Helvetica-Bold')
           .text(`Search: "${term}" (${termResults.length} results)`, { underline: true });
        
        doc.moveDown(0.5);
        
        // Group by store
        const storeGroups = {};
        termResults.forEach(result => {
          if (!storeGroups[result.store]) {
            storeGroups[result.store] = [];
          }
          storeGroups[result.store].push(result);
        });
        
        Object.entries(storeGroups).forEach(([store, items]) => {
          doc.fontSize(14)
             .font('Helvetica-Bold')
             .text(store);
          
          doc.fontSize(11)
             .font('Helvetica');
          
          // FIXED: Only show Quality Code and Stock (Column E)
          items.forEach(item => {
            doc.text(`  ${item.qualityCode}: ${item.stock}`);
          });
          
          doc.moveDown(0.5);
        });
        
        doc.moveDown();
      }
    });

    // Footer
    if (totalResults === 0) {
      doc.fontSize(14)
         .font('Helvetica')
         .text('No matching results found.', { align: 'center' });
      doc.text('Try different search terms or contact support.', { align: 'center' });
    } else {
      doc.fontSize(10)
         .font('Helvetica')
         .text(`Total Results: ${totalResults}`, { align: 'right' });
    }

    doc.end();

    // Wait for PDF to be created
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

// Send file via WhatsApp API
async function sendWhatsAppFile(to, filepath, filename, productId, phoneId) {
  try {
    console.log(`Would send file ${filename} to ${to}`);
    console.log(`File location: ${filepath}`);
    
    // For now, we'll send a message with download instructions
    const fileSize = fs.statSync(filepath).size;
    const fileSizeKB = Math.round(fileSize / 1024);
    
    const message = `📄 *Stock Results Generated*

📊 File: ${filename}
📏 Size: ${fileSizeKB} KB

🔗 Your detailed stock results have been generated as a PDF.

*To download:*
Contact system administrator with this reference: ${filename}

_Type */menu* for main menu._`;
    
    await sendWhatsAppMessage(to, message, productId, phoneId);
    
  } catch (error) {
    console.error('Error sending file:', error);
    throw error;
  }
}

// FIXED: Smart Stock Query with Alphanumeric Support and Clean Output
async function processSmartStockQuery(from, searchTerms, productId, phoneId) {
  try {
    console.log('Processing SMART stock query for:', from);
    console.log('Search terms:', searchTerms);
    
    // FIXED: Validate search terms (minimum 5 characters, alphanumeric allowed)
    const validTerms = searchTerms.filter(term => {
      const cleanTerm = term.trim();
      return cleanTerm.length >= 5; // Accept any 5+ character string (alphanumeric)
    });
    
    if (validTerms.length === 0) {
      await sendWhatsAppMessage(from, `❌ *Invalid Search*

Please provide at least 5 characters for searching.

*Examples:*
- 11010 (finds 11010088471-001)
- ABC12 (finds ABC123456789)
- 88471 (finds 11010088471-001)

_Type */menu* for main menu._`, productId, phoneId);
      return;
    }
    
    await sendWhatsAppMessage(from, `🔍 *Smart Stock Search*

Searching for: ${validTerms.join(', ')}

⏳ Please wait while I search all stock sheets...`, productId, phoneId);

    // Perform smart search
    const searchResults = await searchStockWithPartialMatch(validTerms);
    const permittedStores = await getUserPermittedStores(from);
    
    console.log(`Smart search completed for ${from}`);
    
    // Count total results
    let totalResults = 0;
    validTerms.forEach(term => {
      totalResults += (searchResults[term] || []).length;
    });
    
    console.log(`Total results found: ${totalResults}`);
    
    if (totalResults === 0) {
      // No results found
      const noResultsMessage = `❌ *No Results Found*

No stock items found containing:
${validTerms.map(term => `• ${term}`).join('\n')}

*Try:*
- Different search combinations
- Shorter terms (5+ characters)
- Both letters and numbers work

_Type */menu* for main menu._`;
      
      await sendWhatsAppMessage(from, noResultsMessage, productId, phoneId);
      return;
    }
    
    // Decide: WhatsApp message vs PDF
    if (totalResults <= 15) {
      // SHORT LIST: Send as WhatsApp message - FIXED: Only show code and stock
      let responseMessage = `🎯 *Smart Search Results*\n\n`;
      
      validTerms.forEach(term => {
        const termResults = searchResults[term] || [];
        if (termResults.length > 0) {
          responseMessage += `*"${term}" (${termResults.length} found)*\n`;
          
          // FIXED: Only show Quality Code and Stock Value
          termResults.forEach(result => {
            responseMessage += `${result.qualityCode}: ${result.stock}\n`;
          });
          responseMessage += `\n`;
        }
      });
      
      // Add ordering options if user has permissions
      if (permittedStores.length > 0) {
        responseMessage += `📋 *Place Orders*\n\n`;
        
        if (permittedStores.length === 1) {
          const cleanPhone = from.replace(/^\+/, '');
          const formUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url&entry.740712049=${encodeURIComponent(cleanPhone)}&store=${encodeURIComponent(permittedStores[0])}`;
          responseMessage += `*Your Store:* ${permittedStores}\n${formUrl}\n\n`;
        } else {
          responseMessage += `*Your Stores:*\n`;
          permittedStores.forEach((store, index) => {
            responseMessage += `${index + 1}. ${store}\n`;
          });
          responseMessage += `\nReply with store number to get order form.\n\n`;
        }
      }
      
      responseMessage += `_Type */menu* for main menu._`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      
    } else {
      // LONG LIST: Generate PDF
      try {
        const pdfResult = await generateStockPDF(searchResults, validTerms, from);
        
        const summaryMessage = `📊 *Large Results Found*

🔍 Search: ${validTerms.join(', ')}
📈 Total Results: ${totalResults} items
📄 PDF Generated: ${pdfResult.filename}

*Results are too long for WhatsApp*
📋 Detailed PDF report has been generated.`;
        
        await sendWhatsAppMessage(from, summaryMessage, productId, phoneId);
        
        // Send the PDF file
        await sendWhatsAppFile(from, pdfResult.filepath, pdfResult.filename, productId, phoneId);
        
      } catch (pdfError) {
        console.error('PDF generation failed:', pdfError);
        await sendWhatsAppMessage(from, `❌ *Error Generating PDF*\n\nFound ${totalResults} results but couldn't generate PDF.\nPlease contact support.\n\n_Type */menu* for main menu._`, productId, phoneId);
      }
    }
    
  } catch (error) {
    console.error('Error in smart stock query:', error);
    await sendWhatsAppMessage(from, `❌ *Search Error*\n\nUnable to complete search.\nPlease try again later.\n\n_Type */menu* for main menu._`, productId, phoneId);
  }
}

// Debug permission sheet with proper API handling
async function debugPermissionSheet(phoneNumber) {
  try {
    console.log(`DEBUG: Checking permissions for phone: ${phoneNumber}`);
    
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: STORE_PERMISSION_SHEET_ID,
      range: 'store permission!A:B',
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const rows = response.data.values;
    console.log('Raw data from permission sheet:');
    console.log(JSON.stringify(rows, null, 2));
    
    if (!rows || rows.length === 0) {
      console.log('No data found in permission sheet');
      return;
    }

    const cleanPhone = phoneNumber.replace(/^\+91|^91|^0/, '');
    console.log(`Looking for cleaned phone: "${cleanPhone}"`);
    
    console.log('\nAnalyzing all permission entries:');
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 2) {
        console.log(`Row ${i + 1}: INCOMPLETE ROW - ${JSON.stringify(row)}`);
        continue;
      }
      
      const columnA = row[0] ? row[0].toString().trim() : '';
      const columnB = row[2] ? row[2].toString().trim() : '';
      
      console.log(`Row ${i + 1}:`);
      console.log(`  Column A (raw): "${columnA}"`);
      console.log(`  Column B (raw): "${columnB}"`);
      
      let extractedPhone = '';
      let extractedStore = '';
      
      if (columnA.includes(',')) {
        console.log(`  Malformed data detected in Column A`);
        const parts = columnA.split(',');
        extractedPhone = parts[0].trim();
        extractedStore = columnB || (parts[2] ? parts[2].trim() : '');
      } else {
        extractedPhone = columnA;
        extractedStore = columnB;
      }
      
      const cleanExtracted = extractedPhone.replace(/^\+91|^91|^0/, '');
      const isMatch = cleanExtracted === cleanPhone;
      
      console.log(`  Extracted Phone: "${extractedPhone}" (cleaned: "${cleanExtracted}")`);
      console.log(`  Extracted Store: "${extractedStore}"`);
      console.log(`  ${isMatch ? 'MATCH!' : 'No match'}`);
    }
    
  } catch (error) {
    console.error('Error reading permission sheet:', error);
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

  // ENHANCED: Handle shortcuts and commands with greetings
  const lowerMessage = trimmedMessage.toLowerCase();

  // DEBUG: Permission testing
  if (trimmedMessage.startsWith('DEBUG:')) {
    const phoneToDebug = trimmedMessage.replace('DEBUG:', '').trim() || from;
    await debugPermissionSheet(phoneToDebug);
    await sendWhatsAppMessage(from, `Debug completed for: ${phoneToDebug}\n\nCheck server logs for detailed permission analysis.\n\nType */menu* for main menu.`, productId, phoneId);
    return res.sendStatus(200);
  }

  // NEW: Temporary greeting debug
  if (lowerMessage === '/debuggreet') {
    const greeting = await getUserGreeting(from);
    const message = greeting 
      ? `Found: ${greeting.salutation} ${greeting.name} - ${greeting.greetings}`
      : 'No greeting found';
    
    await sendWhatsAppMessage(from, `Greeting debug: ${message}`, productId, phoneId);
    return res.sendStatus(200);
  }

  // NEW: Smart shortcuts with greetings
  if (lowerMessage === '/menu' || trimmedMessage === '/') {
    console.log('Main menu command received');
    userStates[from] = { currentMenu: 'main' };
    
    // Get user greeting
    const greeting = await getUserGreeting(from);
    
    const mainMenu = `*MAIN MENU*

Please select an option:

1. Ticket
2. Order Query  
3. Stock Query
4. Document

*SHORTCUTS:*
/stock - Direct Stock Query
/shirting - Shirting Orders
/jacket - Jacket Orders  
/trouser - Trouser Orders

_Type the number or use shortcuts..._`;
    
    const finalMessage = formatGreetingMessage(greeting, mainMenu);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // ENHANCED: Direct Stock Query shortcut with greeting and smart search
  if (lowerMessage === '/stock') {
    console.log('Direct stock query shortcut used');
    userStates[from] = { currentMenu: 'smart_stock_query' };
    
    // Get user greeting
    const greeting = await getUserGreeting(from);
    
    const stockQueryPrompt = `*SMART STOCK QUERY* 🔍

Enter any 5+ character code (letters/numbers):

*Examples:*
• 11010 → finds 11010088471-001
• ABC12 → finds ABC123456-XYZ
• 88471 → finds 11010088471-001

*Multiple searches:* Separate with commas
*Example:* 11010, ABC12, 88471

_Smart search finds partial matches!_

Type your search terms below:`;
    
    const finalMessage = formatGreetingMessage(greeting, stockQueryPrompt);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // NEW: Direct Order Query shortcuts with greeting
  if (lowerMessage === '/shirting') {
    console.log('Direct shirting order query shortcut used');
    userStates[from] = { currentMenu: 'order_number_input', category: 'Shirting' };
    
    // Get user greeting
    const greeting = await getUserGreeting(from);
    
    const shirtingQuery = `*SHIRTING ORDER QUERY*

Please enter your Order Number(s):

*Single order:* ABC123
*Multiple orders:* ABC123, DEF456, GHI789

_Type your order numbers below:_`;
    
    const finalMessage = formatGreetingMessage(greeting, shirtingQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  if (lowerMessage === '/jacket') {
    console.log('Direct jacket order query shortcut used');
    userStates[from] = { currentMenu: 'order_number_input', category: 'Jacket' };
    
    // Get user greeting
    const greeting = await getUserGreeting(from);
    
    const jacketQuery = `*JACKET ORDER QUERY*

Please enter your Order Number(s):

*Single order:* ABC123
*Multiple orders:* ABC123, DEF456, GHI789

_Type your order numbers below:_`;
    
    const finalMessage = formatGreetingMessage(greeting, jacketQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  if (lowerMessage === '/trouser') {
    console.log('Direct trouser order query shortcut used');
    userStates[from] = { currentMenu: 'order_number_input', category: 'Trouser' };
    
    // Get user greeting
    const greeting = await getUserGreeting(from);
    
    const trouserQuery = `*TROUSER ORDER QUERY*

Please enter your Order Number(s):

*Single order:* ABC123
*Multiple orders:* ABC123, DEF456, GHI789

_Type your order numbers below:_`;
    
    const finalMessage = formatGreetingMessage(greeting, trouserQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle menu selections (existing logic)
  if (userStates[from] && userStates[from].currentMenu === 'main') {
    if (trimmedMessage === '1') {
      const ticketMenu = `*TICKET OPTIONS*

Click the links below to access forms directly:

*HELP TICKET*
${links.helpTicket}

*LEAVE FORM*
${links.leave}

*DELEGATION*
${links.delegation}

_Type */menu* to return to main menu._`;
      
      await sendWhatsAppMessage(from, ticketMenu, productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2') {
      userStates[from].currentMenu = 'order_query';
      const orderQueryMenu = `*ORDER QUERY*

Please select the product category:

1. Shirting
2. Jacket  
3. Trouser

_Type the number to continue..._`;
      
      await sendWhatsAppMessage(from, orderQueryMenu, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
      userStates[from] = { currentMenu: 'smart_stock_query' };
      
      const stockQueryPrompt = `*SMART STOCK QUERY* 🔍

Enter any 5+ character code (letters/numbers):

*Examples:*
• 11010 → finds 11010088471-001
• ABC12 → finds ABC123456-XYZ
• 88471 → finds 11010088471-001

*Multiple searches:* Separate with commas
*Example:* 11010, ABC12, 88471

_Smart search finds partial matches!_

Type your search terms below:`;
      
      await sendWhatsAppMessage(from, stockQueryPrompt, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '4') {
      await sendWhatsAppMessage(from, '*DOCUMENT*\n\nThis feature is coming soon!\n\nType */menu* to return to main menu.', productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }

    await sendWhatsAppMessage(from, 'Invalid option. Please select 1, 2, 3, or 4.\n\nType */menu* to see the main menu again.', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle order query category selection
  if (userStates[from] && userStates[from].currentMenu === 'order_query') {
    if (trimmedMessage === '1' || lowerMessage === '/shirting') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Shirting' };
      await sendWhatsAppMessage(from, `*SHIRTING ORDER QUERY*\n\nPlease enter your Order Number(s):\n\n*Single order:* ABC123\n*Multiple orders:* ABC123, DEF456, GHI789\n\n_Type your order numbers below:_`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2' || lowerMessage === '/jacket') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Jacket' };
      await sendWhatsAppMessage(from, `*JACKET ORDER QUERY*\n\nPlease enter your Order Number(s):\n\n*Single order:* ABC123\n*Multiple orders:* ABC123, DEF456, GHI789\n\n_Type your order numbers below:_`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3' || lowerMessage === '/trouser') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Trouser' };
      await sendWhatsAppMessage(from, `*TROUSER ORDER QUERY*\n\nPlease enter your Order Number(s):\n\n*Single order:* ABC123\n*Multiple orders:* ABC123, DEF456, GHI789\n\n_Type your order numbers below:_`, productId, phoneId);
      return res.sendStatus(200);
    }

    await sendWhatsAppMessage(from, 'Invalid option. Please select 1, 2, or 3.\n\nType */menu* to return to main menu.', productId, phoneId);
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

  // FIXED: Handle smart stock query input (alphanumeric support)
  if (userStates[from] && userStates[from].currentMenu === 'smart_stock_query') {
    if (trimmedMessage !== '/menu') {
      const searchTerms = trimmedMessage.split(',').map(q => q.trim()).filter(q => q.length > 0);
      await processSmartStockQuery(from, searchTerms, productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }
  }

  // Handle multiple order selection
  if (userStates[from] && userStates[from].currentMenu === 'multiple_order_selection') {
    if (trimmedMessage !== '/menu') {
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
    
    await sendWhatsAppMessage(from, `*Checking ${category} orders...*\n\nPlease wait while I search for your order status.`, productId, phoneId);

    let responseMessage = `*${category.toUpperCase()} ORDER STATUS*\n\n`;
    
    for (const orderNum of orderNumbers) {
      console.log(`Searching for order: ${orderNum}`);
      
      const orderStatus = await searchOrderStatus(orderNum, category);
      
      responseMessage += `*Order: ${orderNum}*\n`;
      responseMessage += `${orderStatus.message}\n\n`;
    }
    
    responseMessage += `_Type */menu* to return to main menu._`;
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
  } catch (error) {
    console.error('Error processing order query:', error);
    await sendWhatsAppMessage(from, 'Error checking orders\n\nPlease try again later.\n\nType */menu* to return to main menu.', productId, phoneId);
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
      
      if (row.toString().trim() === orderNumber.trim()) {
        console.log(`Order ${orderNumber} found in FMS sheet at row ${i + 1}`);
        
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
      
      if (row[3].toString().trim() === orderNumber.trim()) {
        console.log(`Order ${orderNumber} found in completed sheet at row ${i + 1}`);
        
        // Get dispatch date from column CH (index 87)
        const dispatchDate = row[87] ? row[87].toString().trim() : 'Date not available';
        
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

// Check ALL production stages for last completed one
function checkProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let lastCompletedStageIndex = -1;
    let hasAnyStage = false;

    console.log('Checking ALL production stages for last completed...');

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
        console.log(`Stage ${stage.name} completed: ${cellValue} (Index: ${i})`);
      } else {
        console.log(`Stage ${stage.name} not completed`);
      }
    }

    console.log(`Last completed stage: ${lastCompletedStage ? lastCompletedStage.name : 'None'} (Index: ${lastCompletedStageIndex})`);

    // Generate status message based on findings
    if (!hasAnyStage) {
      return { message: 'Order is currently under process' };
    }

    // Check if it's the final stage (Dispatch HO)
    if (lastCompletedStage && lastCompletedStage.name === 'Dispatch (HO)') {
      // Get dispatch date from column CH
      const dispatchDateIndex = columnToIndex(lastCompletedStage.dispatchDateColumn);
      const dispatchDate = row[dispatchDateIndex] ? row[dispatchDateIndex].toString().trim() : 'Date not available';
      return { message: `Order has been dispatched from HO on ${dispatchDate}` };
    }

    // For any other completed stage, show current stage and next stage
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

// Convert column letter to index
function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

// Get user permitted stores with proper API handling
async function getUserPermittedStores(phoneNumber) {
  try {
    console.log(`Getting permitted stores for phone: ${phoneNumber}`);
    
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    // Use explicit sheet name and proper API options
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: STORE_PERMISSION_SHEET_ID,
      range: 'store permission!A:B',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log('No data found in Store Permission sheet');
      return [];
    }

    console.log(`Found ${rows.length} total rows in permission sheet`);
    console.log('Raw rows data:', JSON.stringify(rows.slice(0, 5), null, 2));
    
    const permittedStores = [];
    
    // Clean incoming phone number multiple ways
    const phoneVariations = [
      phoneNumber,
      phoneNumber.replace(/^\+91/, ''),
      phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''),
      phoneNumber.replace(/^0/, ''),
      phoneNumber.replace(/[\s\-\(\)]/g, ''),
    ];
    
    console.log(`Trying phone variations: ${phoneVariations.join(', ')}`);
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 2) {
        console.log(`Row ${i + 1}: Skipping incomplete row: ${JSON.stringify(row)}`);
        continue;
      }
      
      let sheetContact = '';
      let sheetStore = '';
      
      const columnAValue = row[0] ? row.toString().trim() : '';
      const columnBValue = row[2] ? row[2].toString().trim() : '';
      
      if (columnAValue.includes(',')) {
        console.log(`Row ${i + 1}: Detected malformed data in Column A: "${columnAValue}"`);
        const parts = columnAValue.split(',');
        sheetContact = parts[0].trim();
        sheetStore = columnBValue || (parts[2] ? parts[2].trim() : '');
      } else {
        sheetContact = columnAValue;
        sheetStore = columnBValue;
      }
      
      const sheetContactVariations = [
        sheetContact,
        sheetContact.replace(/^\+91/, ''),
        sheetContact.replace(/^\+/, ''),
        sheetContact.replace(/^91/, ''),
        sheetContact.replace(/^0/, ''),
        sheetContact.replace(/[\s\-\(\)]/g, ''),
      ];
      
      console.log(`Row ${i + 1}: Contact="${sheetContact}" → Store="${sheetStore}"`);
      console.log(`  Contact variations: ${sheetContactVariations.join(', ')}`);
      
      let isMatch = false;
      for (const phoneVar of phoneVariations) {
        for (const sheetVar of sheetContactVariations) {
          if (phoneVar === sheetVar) {
            console.log(`  MATCH FOUND! "${phoneVar}" === "${sheetVar}"`);
            isMatch = true;
            break;
          }
        }
        if (isMatch) break;
      }
      
      if (isMatch) {
        permittedStores.push(sheetStore);
        console.log(`  Added permitted store: ${sheetStore}`);
      }
    }
    
    console.log(`Final result: Found ${permittedStores.length} permitted stores for ${phoneNumber}:`, permittedStores);
    return permittedStores;
    
  } catch (error) {
    console.error('Error getting permitted stores:', error);
    return [];
  }
}

// Handle multiple order selection
async function handleMultipleOrderSelectionWithHiddenField(from, userInput, productId, phoneId) {
  try {
    const userState = userStates[from];
    if (!userState) {
      await sendWhatsAppMessage(from, 'Session expired. Start over: /stock or /menu', productId, phoneId);
      return;
    }
    
    if (/^\d+$/.test(userInput.trim())) {
      const storeIndex = parseInt(userInput) - 1;
      const selectedStore = userState.permittedStores[storeIndex];
      
      if (selectedStore) {
        await createSingleStoreForm(from, selectedStore, userState.qualities, productId, phoneId);
      } else {
        await sendWhatsAppMessage(from, `Invalid store number`, productId, phoneId);
      }
    } else {
      const combinations = parseMultipleOrderInput(userInput, userState);
      
      if (combinations.length > 0) {
        await createMultipleStoreForms(from, combinations, productId, phoneId);
      } else {
        await sendWhatsAppMessage(from, 'Invalid format. Try: Quality-StoreNumber, Quality-StoreNumber', productId, phoneId);
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
        const storeIndex = parseInt(match[2]) - 1;
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
    
    let confirmationMessage = `*${selectedStore}*\n\n`;
    confirmationMessage += `${formUrl}\n\n`;
    confirmationMessage += `_Type */menu* for main menu._`;
    
    await sendWhatsAppMessage(from, confirmationMessage, productId, phoneId);
    
    userStates[from] = { currentMenu: 'completed' };
    
  } catch (error) {
    console.error('Error creating single store form:', error);
  }
}

async function createMultipleStoreForms(from, combinations, productId, phoneId) {
  try {
    const cleanPhone = from.replace(/^\+/, '');
    let responseMessage = `*MULTIPLE ORDER FORMS*\n\n`;
    
    const storeGroups = {};
    combinations.forEach(combo => {
      if (!storeGroups[combo.store]) {
        storeGroups[combo.store] = [];
      }
      storeGroups[combo.store].push(combo.quality);
    });
    
    Object.entries(storeGroups).forEach(([store, qualities]) => {
      const formUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url` +
        `&entry.740712049=${encodeURIComponent(cleanPhone)}` +
        `&store=${encodeURIComponent(store)}`;
      
      responseMessage += `*${store}*\n`;
      responseMessage += `${qualities.join(', ')}\n`;
      responseMessage += `${formUrl}\n\n`;
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
  console.log(`WhatsApp Bot running on port ${PORT}`);
  console.log('🚀 Bot ready with FIXED SMART STOCK SEARCH!');
  console.log('✨ Fixed Features:');
  console.log('🔍 Alphanumeric partial matching (5+ characters: letters/numbers)');
  console.log('📄 Clean output showing only Quality Code and Stock (Column E)');
  console.log('⚡ Fast substring search across all stock sheets');
  console.log('📊 Smart output: WhatsApp text vs PDF based on result count');
  console.log('🎯 User-friendly examples for alphanumeric search');
  console.log('Available shortcuts: /menu, /stock, /shirting, /jacket, /trouser');
});
