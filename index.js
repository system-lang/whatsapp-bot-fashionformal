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
    
    console.log(`Download request for: ${filename}`);
    
    if (fs.existsSync(filepath)) {
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.download(filepath, filename, (err) => {
        if (err) {
          console.error('Download error:', err);
        } else {
          console.log(`Successfully served: ${filename}`);
          setTimeout(() => {
            try {
              fs.unlinkSync(filepath);
              console.log(`Cleaned up: ${filename}`);
            } catch (cleanupErr) {
              console.log('Cleanup error:', cleanupErr.message);
            }
          }, 300000);
        }
      });
    } else {
      console.log(`File not found: ${filepath}`);
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

// FIXED: Get user greeting from Greetings sheet with better debugging
async function getUserGreeting(phoneNumber) {
  try {
    console.log(`GREETING DEBUG: Getting greeting for phone: ${phoneNumber}`);
    
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    let rows = null;
    
    try {
      const metaResponse = await sheets.spreadsheets.get({
        spreadsheetId: GREETINGS_SHEET_ID,
      });
      
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
        console.log(`GREETING DEBUG: Found greetings sheet: "${sheetName}"`);
        
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: GREETINGS_SHEET_ID,
          range: `'${sheetName}'!A:D`,
          valueRenderOption: 'UNFORMATTED_VALUE'
        });
        
        rows = response.data.values;
        console.log(`GREETING DEBUG: Successfully read ${rows ? rows.length : 0} rows from greetings sheet`);
      }
    } catch (metaError) {
      console.log('GREETING DEBUG: Metadata approach failed, trying alternative ranges...');
      const attempts = ['Greetings!A:D', "'Greetings'!A:D", 'Sheet2!A:D', 'A:D'];
      
      for (const range of attempts) {
        try {
          console.log(`GREETING DEBUG: Trying range: ${range}`);
          const response = await sheets.spreadsheets.values.get({
            spreadsheetId: GREETINGS_SHEET_ID,
            range: range,
            valueRenderOption: 'UNFORMATTED_VALUE'
          });
          
          if (response.data.values && response.data.values.length > 0) {
            rows = response.data.values;
            console.log(`GREETING DEBUG: Successfully accessed greetings with range: ${range}`);
            break;
          }
        } catch (rangeError) {
          continue;
        }
      }
    }
    
    if (!rows || rows.length === 0) {
      console.log('GREETING DEBUG: No data found in any Greetings sheet attempt');
      return null;
    }

    console.log('GREETING DEBUG: Greetings sheet raw data:', JSON.stringify(rows.slice(0, 3), null, 2));

    const phoneVariations = Array.from(new Set([
      phoneNumber,
      phoneNumber.replace(/^\+91/, ''),
      phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''),
      phoneNumber.replace(/^0/, ''),
      phoneNumber.replace(/[\s\-\(\)]/g, ''),
      phoneNumber.slice(-10)
    ])).filter(p => p && p.length >= 10);
    
    console.log(`GREETING DEBUG: Looking for greeting with phone variations: ${phoneVariations.join(', ')}`);
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) {
        console.log(`GREETING DEBUG: Row ${i + 1}: Skipping empty greeting row`);
        continue;
      }
      
      let sheetContact, name, salutation, greetings;
      
      if (row.length >= 4 && row[1] && row[2] && row[3] && !row.toString().includes(',')) {
        // Normal format
        const contactCell = row;
        if (typeof contactCell === 'number') {
          sheetContact = contactCell.toString();
          if (sheetContact.includes('e+') || sheetContact.includes('E+')) {
            sheetContact = contactCell.toFixed(0);
          }
        } else {
          sheetContact = contactCell ? contactCell.toString().trim() : '';
        }
        name = row[11].toString().trim();
        salutation = row[12].toString().trim(); 
        greetings = row.toString().trim();
        console.log(`GREETING DEBUG: Row ${i + 1}: NORMAL - Contact="${sheetContact}", Name="${name}", Salutation="${salutation}", Greetings="${greetings}"`);
      } else if (row && row.toString().includes(',')) {
        // Comma-separated format
        const parts = row.toString().split(',');
        if (parts.length >= 4) {
          sheetContact = parts.trim();
          name = parts[11].trim();
          salutation = parts[12].trim(); 
          greetings = parts.trim();
          console.log(`GREETING DEBUG: Row ${i + 1}: COMMA - Contact="${sheetContact}", Name="${name}", Salutation="${salutation}", Greetings="${greetings}"`);
        } else {
          console.log(`GREETING DEBUG: Row ${i + 1}: Insufficient parts: ${JSON.stringify(parts)}`);
          continue;
        }
      } else {
        console.log(`GREETING DEBUG: Row ${i + 1}: Unrecognized format: ${JSON.stringify(row)}`);
        continue;
      }
      
      // Check for match
      const sheetContactVariations = Array.from(new Set([
        sheetContact,
        sheetContact.replace(/^\+91/, ''),
        sheetContact.replace(/^\+/, ''),
        sheetContact.replace(/^91/, ''),
        sheetContact.replace(/^0/, ''),
        sheetContact.replace(/[\s\-\(\)]/g, ''),
        sheetContact.slice(-10)
      ])).filter(s => s && s.length >= 10);
      
      let isMatch = false;
      for (const phoneVar of phoneVariations) {
        if (sheetContactVariations.includes(phoneVar)) {
          console.log(`GREETING DEBUG: MATCH found! "${phoneVar}"`);
          isMatch = true;
          break;
        }
      }
      
      if (isMatch) {
        console.log(`GREETING DEBUG: SUCCESS! Found greeting for ${phoneNumber}: ${salutation} ${name} - ${greetings}`);
        return { name, salutation, greetings };
      }
    }
    
    console.log(`GREETING DEBUG: No greeting found for ${phoneNumber} after checking all rows`);
    return null;
    
  } catch (error) {
    console.error('GREETING DEBUG: Error getting user greeting:', error);
    return null;
  }
}

// FIXED: Format greeting message with better debugging
function formatGreetingMessage(greeting, mainMessage) {
  console.log(`FORMAT DEBUG: Greeting object:`, greeting);
  console.log(`FORMAT DEBUG: Main message length:`, mainMessage ? mainMessage.length : 'null');
  
  if (!greeting) {
    console.log(`FORMAT DEBUG: No greeting found, returning main message only`);
    return mainMessage;
  }
  
  const formattedMessage = `${greeting.salutation} ${greeting.name}\n\n${greeting.greetings}\n\n${mainMessage}`;
  console.log(`FORMAT DEBUG: Formatted message:`, formattedMessage.substring(0, 100) + '...');
  return formattedMessage;
}

// FIXED: Enhanced Smart Stock Search with proper column parsing
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

    console.log(`Found ${folderFiles.data.files.length} stock files for smart search`);

    for (const file of folderFiles.data.files) {
      console.log(`Smart searching in: ${file.name}`);
      
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
        
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row[0]) continue;
          
          const qualityCode = row.toString().trim();
          
          // FIXED: Parse stock value correctly - handle comma-separated data
          let stockValue = '0';
          if (row) {
            const rawStock = row.toString().trim();
            // If the stock value contains commas, get the last non-empty value
            if (rawStock.includes(',')) {
              const stockParts = rawStock.split(',').map(part => part.trim()).filter(part => part !== '');
              stockValue = stockParts.length > 0 ? stockParts[stockParts.length - 1] : '0';
            } else {
              stockValue = rawStock;
            }
          }
          
          console.log(`STOCK DEBUG: Raw row[4]: "${row}" -> Parsed stock: "${stockValue}"`);
          
          searchTerms.forEach(searchTerm => {
            const cleanSearchTerm = searchTerm.trim();
            
            if (cleanSearchTerm.length >= 5 && qualityCode.toUpperCase().includes(cleanSearchTerm.toUpperCase())) {
              console.log(`MATCH: "${cleanSearchTerm}" found in "${qualityCode}" - Stock: ${stockValue}`);
              
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

// Generate PDF showing ONLY Quality Code and Stock (Professional Format)
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

    // Professional Header
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

    // Draw line separator
    doc.moveTo(50, doc.y)
       .lineTo(550, doc.y)
       .stroke();
    doc.moveDown(0.5);

    // Results - ONLY Quality Code and Stock
    let totalResults = 0;
    
    searchTerms.forEach(term => {
      const termResults = searchResults[term] || [];
      totalResults += termResults.length;
      
      if (termResults.length > 0) {
        doc.fontSize(14)
           .font('Helvetica-Bold')
           .text(`Search Term: "${term}" (${termResults.length} results found)`);
        
        doc.moveDown(0.3);
        
        // Professional table header
        doc.fontSize(10)
           .font('Helvetica-Bold')
           .text('Quality Code', 50, doc.y, { width: 300, continued: true })
           .text('Stock Quantity', 350, doc.y);
        
        doc.moveDown(0.2);
        
        // Draw header line
        doc.moveTo(50, doc.y)
           .lineTo(550, doc.y)
           .stroke();
        doc.moveDown(0.3);
        
        // ONLY show Quality Code and Stock - NO other columns
        doc.fontSize(9)
           .font('Helvetica');
        
        termResults.forEach(item => {
          doc.text(item.qualityCode, 50, doc.y, { width: 300, continued: true })
             .text(item.stock, 350, doc.y);
          doc.moveDown(0.2);
        });
        
        doc.moveDown(0.5);
      }
    });

    // Footer
    if (totalResults === 0) {
      doc.fontSize(12)
         .font('Helvetica')
         .text('No matching results found.', { align: 'center' });
      doc.text('Please try different search terms.', { align: 'center' });
    } else {
      doc.fontSize(8)
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

// Send file via Railway file serving endpoint
async function sendWhatsAppFile(to, filepath, filename, productId, phoneId) {
  try {
    console.log(`Preparing download link for: ${filename}`);
    
    const fileStats = fs.statSync(filepath);
    const fileSizeKB = Math.round(fileStats.size / 1024);
    
    const baseUrl = 'https://whatsapp-bot-fashionformal-production.up.railway.app';
    const downloadUrl = `${baseUrl}/download/${filename}`;
    
    const message = `*Stock Results Generated*

File: ${filename}
Size: ${fileSizeKB} KB

Download your PDF:
${downloadUrl}

Click the link above to download
Works on mobile and desktop
Link expires in 5 minutes

Type /menu for main menu`;
    
    await sendWhatsAppMessage(to, message, productId, phoneId);
    
  } catch (error) {
    console.error('Error creating download link:', error);
    
    const fallbackMessage = `PDF Generated Successfully

Your stock results have been generated.

Contact support for file access:
Reference: ${filename}

Type /menu for main menu`;
    
    await sendWhatsAppMessage(to, fallbackMessage, productId, phoneId);
  }
}

// FIXED: Get user permitted stores with proper store name parsing
async function getUserPermittedStores(phoneNumber) {
  try {
    console.log(`STORE DEBUG: Getting permitted stores for phone: ${phoneNumber}`);
    
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: STORE_PERMISSION_SHEET_ID,
      range: 'store permission!A:B',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log('STORE DEBUG: No data found in Store Permission sheet');
      return [];
    }

    console.log(`STORE DEBUG: Found ${rows.length} total rows in permission sheet`);
    console.log('STORE DEBUG: Raw rows data:', JSON.stringify(rows.slice(0, 5), null, 2));
    
    const permittedStores = [];
    
    const phoneVariations = [
      phoneNumber,
      phoneNumber.replace(/^\+91/, ''),
      phoneNumber.replace(/^\+/, ''),
      phoneNumber.replace(/^91/, ''),
      phoneNumber.replace(/^0/, ''),
      phoneNumber.replace(/[\s\-\(\)]/g, ''),
    ];
    
    console.log(`STORE DEBUG: Trying phone variations: ${phoneVariations.join(', ')}`);
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 1) {
        console.log(`STORE DEBUG: Row ${i + 1}: Skipping incomplete row: ${JSON.stringify(row)}`);
        continue;
      }
      
      let sheetContact = '';
      let sheetStore = '';
      
      const columnAValue = row[0] ? row.toString().trim() : '';
      const columnBValue = row[11] ? row[11].toString().trim() : '';
      
      // FIXED: Handle different data formats in store permission sheet
      if (columnAValue.includes(',')) {
        console.log(`STORE DEBUG: Row ${i + 1}: Detected comma-separated data in Column A: "${columnAValue}"`);
        const parts = columnAValue.split(',').map(part => part.trim());
        sheetContact = parts;
        // Try to find store name in the parts or use column B
        if (parts.length > 1 && parts[11]) {
          sheetStore = parts[11];
        } else {
          sheetStore = columnBValue;
        }
      } else {
        sheetContact = columnAValue;
        sheetStore = columnBValue;
      }
      
      // Handle case where contact number is a JavaScript number
      if (typeof sheetContact === 'number') {
        sheetContact = sheetContact.toString();
        if (sheetContact.includes('e+') || sheetContact.includes('E+')) {
          sheetContact = sheetContact.replace(/e\+\d+/gi, '');
        }
      }
      
      const sheetContactVariations = [
        sheetContact,
        sheetContact.replace(/^\+91/, ''),
        sheetContact.replace(/^\+/, ''),
        sheetContact.replace(/^91/, ''),
        sheetContact.replace(/^0/, ''),
        sheetContact.replace(/[\s\-\(\)]/g, ''),
      ];
      
      console.log(`STORE DEBUG: Row ${i + 1}: Contact="${sheetContact}" Store="${sheetStore}"`);
      console.log(`STORE DEBUG: Contact variations: ${sheetContactVariations.join(', ')}`);
      
      let isMatch = false;
      for (const phoneVar of phoneVariations) {
        for (const sheetVar of sheetContactVariations) {
          if (phoneVar === sheetVar) {
            console.log(`STORE DEBUG: MATCH FOUND! "${phoneVar}" === "${sheetVar}"`);
            isMatch = true;
            break;
          }
        }
        if (isMatch) break;
      }
      
      if (isMatch && sheetStore && sheetStore.trim() !== '') {
        permittedStores.push(sheetStore);
        console.log(`STORE DEBUG: Added permitted store: ${sheetStore}`);
      } else if (isMatch) {
        console.log(`STORE DEBUG: Match found but store name is empty for contact: ${sheetContact}`);
      }
    }
    
    console.log(`STORE DEBUG: Final result: Found ${permittedStores.length} permitted stores for ${phoneNumber}:`, permittedStores);
    return permittedStores;
    
  } catch (error) {
    console.error('STORE DEBUG: Error getting permitted stores:', error);
    return [];
  }
}

// FIXED: Smart Stock Query with proper greetings and order forms
async function processSmartStockQuery(from, searchTerms, productId, phoneId) {
  try {
    console.log('Processing SMART stock query for:', from);
    console.log('Search terms:', searchTerms);
    
    // Validate search terms (minimum 5 characters, alphanumeric allowed)
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
    
    // Decide: WhatsApp message vs PDF
    if (totalResults <= 15) {
      // SHORT LIST: Send as WhatsApp message
      let responseMessage = `*Smart Search Results*\n\n`;
      
      validTerms.forEach(term => {
        const termResults = searchResults[term] || [];
        if (termResults.length > 0) {
          responseMessage += `*"${term}" (${termResults.length} found)*\n`;
          
          termResults.forEach(result => {
            responseMessage += `${result.qualityCode}: ${result.stock}\n`;
          });
          responseMessage += `\n`;
        }
      });
      
      // FIXED: Always show ordering section with proper permissions handling
      responseMessage += `*Place Orders*\n\n`;

      if (permittedStores.length === 0) {
        responseMessage += `No store permissions found for ${from}\n`;
        responseMessage += `Contact admin to get ordering access.\n\n`;
      } else if (permittedStores.length === 1) {
        const cleanPhone = from.replace(/^\+/, '');
        const formUrl = `${STATIC_FORM_BASE_URL}?usp=pp_url&entry.740712049=${encodeURIComponent(cleanPhone)}&store=${encodeURIComponent(permittedStores[0])}`;
        responseMessage += `*Your Store:* ${permittedStores}\n${formUrl}\n\n`;
      } else {
        responseMessage += `*Your Stores:*\n`;
        permittedStores.forEach((store, index) => {
          responseMessage += `${index + 1}. ${store}\n`;
        });
        responseMessage += `\nReply with store number to get order form.\n\n`;
        
        // Set up multiple order selection state
        userStates[from] = {
          currentMenu: 'multiple_order_selection',
          permittedStores: permittedStores,
          qualities: validTerms
        };
      }
      
      responseMessage += `Type /menu for main menu`;
      
      await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
      
    } else {
      // LONG LIST: Generate PDF
      try {
        const pdfResult = await generateStockPDF(searchResults, validTerms, from);
        
        const summaryMessage = `*Large Results Found*

Search: ${validTerms.join(', ')}
Total Results: ${totalResults} items
PDF Generated: ${pdfResult.filename}

Results are too long for WhatsApp
Detailed PDF report has been generated`;
        
        await sendWhatsAppMessage(from, summaryMessage, productId, phoneId);
        
        // Send the PDF download link
        await sendWhatsAppFile(from, pdfResult.filepath, pdfResult.filename, productId, phoneId);
        
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

  const lowerMessage = trimmedMessage.toLowerCase();

  // Debug commands
  if (lowerMessage === '/debuggreet') {
    console.log('DEBUG COMMAND: Testing greeting functionality');
    const greeting = await getUserGreeting(from);
    const debugMessage = greeting 
      ? `Found: ${greeting.salutation} ${greeting.name} - ${greeting.greetings}`
      : 'No greeting found';
    
    await sendWhatsAppMessage(from, `Greeting debug: ${debugMessage}`, productId, phoneId);
    return res.sendStatus(200);
  }

  // FIXED: Main menu with greeting and debugging
  if (lowerMessage === '/menu' || trimmedMessage === '/') {
    console.log('Main menu command received - FETCHING GREETING');
    userStates[from] = { currentMenu: 'main' };
    
    const greeting = await getUserGreeting(from);
    console.log('MENU DEBUG: Greeting result:', greeting);
    
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

Type the number or use shortcuts`;
    
    const finalMessage = formatGreetingMessage(greeting, mainMenu);
    console.log('MENU DEBUG: Final message length:', finalMessage.length);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // FIXED: Direct Stock Query shortcut with greeting and debugging
  if (lowerMessage === '/stock') {
    console.log('Direct stock query shortcut used - FETCHING GREETING');
    userStates[from] = { currentMenu: 'smart_stock_query' };
    
    const greeting = await getUserGreeting(from);
    console.log('STOCK DEBUG: Greeting result:', greeting);
    
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
    console.log('STOCK DEBUG: Final message length:', finalMessage.length);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Order Query shortcuts with greeting
  if (lowerMessage === '/shirting') {
    console.log('Direct shirting order query shortcut used - FETCHING GREETING');
    userStates[from] = { currentMenu: 'order_number_input', category: 'Shirting' };
    
    const greeting = await getUserGreeting(from);
    
    const shirtingQuery = `*SHIRTING ORDER QUERY*

Please enter your Order Number(s):

Single order: ABC123
Multiple orders: ABC123, DEF456, GHI789

Type your order numbers below:`;
    
    const finalMessage = formatGreetingMessage(greeting, shirtingQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  if (lowerMessage === '/jacket') {
    console.log('Direct jacket order query shortcut used - FETCHING GREETING');
    userStates[from] = { currentMenu: 'order_number_input', category: 'Jacket' };
    
    const greeting = await getUserGreeting(from);
    
    const jacketQuery = `*JACKET ORDER QUERY*

Please enter your Order Number(s):

Single order: ABC123
Multiple orders: ABC123, DEF456, GHI789

Type your order numbers below:`;
    
    const finalMessage = formatGreetingMessage(greeting, jacketQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  if (lowerMessage === '/trouser') {
    console.log('Direct trouser order query shortcut used - FETCHING GREETING');
    userStates[from] = { currentMenu: 'order_number_input', category: 'Trouser' };
    
    const greeting = await getUserGreeting(from);
    
    const trouserQuery = `*TROUSER ORDER QUERY*

Please enter your Order Number(s):

Single order: ABC123
Multiple orders: ABC123, DEF456, GHI789

Type your order numbers below:`;
    
    const finalMessage = formatGreetingMessage(greeting, trouserQuery);
    await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle menu selections
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

Type /menu to return to main menu`;
      
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

Type the number to continue`;
      
      await sendWhatsAppMessage(from, orderQueryMenu, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
      console.log('Menu option 3 selected - FETCHING GREETING');
      userStates[from] = { currentMenu: 'smart_stock_query' };
      
      const greeting = await getUserGreeting(from);
      console.log('MENU3 DEBUG: Greeting result:', greeting);
      
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
      console.log('MENU3 DEBUG: Final message length:', finalMessage.length);
      await sendWhatsAppMessage(from, finalMessage, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '4') {
      await sendWhatsAppMessage(from, '*DOCUMENT*\n\nThis feature is coming soon!\n\nType /menu to return to main menu', productId, phoneId);
      userStates[from].currentMenu = 'completed';
      return res.sendStatus(200);
    }

    await sendWhatsAppMessage(from, 'Invalid option. Please select 1, 2, 3, or 4.\n\nType /menu to see the main menu again', productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle order query category selection
  if (userStates[from] && userStates[from].currentMenu === 'order_query') {
    if (trimmedMessage === '1' || lowerMessage === '/shirting') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Shirting' };
      await sendWhatsAppMessage(from, `*SHIRTING ORDER QUERY*\n\nPlease enter your Order Number(s):\n\nSingle order: ABC123\nMultiple orders: ABC123, DEF456, GHI789\n\nType your order numbers below:`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2' || lowerMessage === '/jacket') {
      userStates[from] = { currentMenu: 'order_number_input', category: 'Jacket' };
      await sendWhatsAppMessage(from, `*JACKET ORDER QUERY*\n\nPlease enter your Order Number(s):\n\nSingle order: ABC123\nMultiple orders: ABC123, DEF456, GHI789\n\nType your order numbers below:`, productId, phoneId);
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3' || lowerMessage === '/trouser') {
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

// [Include all remaining functions with proper debugging]

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
    
    responseMessage += `Type /menu to return to main menu`;
    
    await sendWhatsAppMessage(from, responseMessage, productId, phoneId);
    
  } catch (error) {
    console.error('Error processing order query:', error);
    await sendWhatsAppMessage(from, 'Error checking orders\n\nPlease try again later.\n\nType /menu to return to main menu', productId, phoneId);
  }
}

async function searchOrderStatus(orderNumber, category) {
  try {
    const auth = await getGoogleAuth();
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const drive = google.drive({ version: 'v3', auth: authClient });

    console.log(`Searching ${orderNumber} in live sheet (FMS)`);
    
    const liveSheetResult = await searchInLiveSheet(sheets, orderNumber);
    if (liveSheetResult.found) {
      console.log(`Order ${orderNumber} found in live sheet`);
      return liveSheetResult;
    }

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
      if (!row[3]) continue;
      
      if (row.toString().trim() === orderNumber.trim()) {
        console.log(`Order ${orderNumber} found in completed sheet at row ${i + 1}`);
        
        const dispatchDate = row[13] ? row[13].toString().trim() : 'Date not available';
        
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

function checkProductionStages(row) {
  try {
    let lastCompletedStage = null;
    let hasAnyStage = false;

    console.log('Checking ALL production stages for last completed...');

    for (let i = 0; i < PRODUCTION_STAGES.length; i++) {
      const stage = PRODUCTION_STAGES[i];
      const columnIndex = columnToIndex(stage.column);
      const cellValue = row[columnIndex] ? row[columnIndex].toString().trim() : '';
      
      console.log(`Stage ${stage.name} (Col ${stage.column}): "${cellValue}"`);
      
      if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
        lastCompletedStage = stage;
        hasAnyStage = true;
        console.log(`Stage ${stage.name} completed: ${cellValue}`);
      } else {
        console.log(`Stage ${stage.name} not completed`);
      }
    }

    console.log(`Last completed stage: ${lastCompletedStage ? lastCompletedStage.name : 'None'}`);

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

function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

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
        const storeIndex = parseInt(match[12]) - 1;
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
    confirmationMessage += `Type /menu for main menu`;
    
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
    
    responseMessage += `Fill each form for your different store orders`;
    
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
  console.log('FINAL DEBUGGING VERSION - All Issues Fixed:');
  console.log('✅ Stock values parsed correctly (handles comma-separated data)');
  console.log('✅ Store permissions parsed correctly (handles various formats)');
  console.log('✅ Greetings working in all stock queries with extensive debugging');
  console.log('✅ Order forms always visible in stock query results');
  console.log('✅ PDF shows ONLY Quality Code and Stock (professional format)');
  console.log('✅ Professional formatting without excessive emojis');
  console.log('✅ Smart partial matching (5+ characters: letters/numbers)');
  console.log('✅ PDF download via Railway file serving');
  console.log('Available shortcuts: /menu, /stock, /shirting, /jacket, /trouser');
  console.log('Debug commands: /debuggreet - Test greeting functionality');
});
