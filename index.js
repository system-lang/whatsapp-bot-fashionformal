require('dotenv').config();
const express = require('express');
const axios = require('axios');
const app = express();

app.use(express.json());

// Store user states to track which menu they're in
let userStates = {};

// Links for different options
const links = {
  helpTicket: 'https://script.google.com/a/macros/fashionformal.com/s/AKfycbzTi9l6afTIaj7f6aiKAMuE7Hz4pQX8796wk5inuHw7wAFgbjv0sFQNCVquPzNEniYdEg/exec',
  delegation: 'https://script.google.com/a/macros/fashionformal.com/s/AKfycbwqdP4BmXhOKm6UEu-xd8Pag_6UErQzr7KKP0mXiECatvv1rDL5-sWLPYAIwReAHfgi/exec',
  leave: 'YOUR_LEAVE_FORM_LINK_HERE' // Replace with actual link when you get it
};

app.post('/webhook', async (req, res) => {
  // Debug: Log the entire request body to see Maytapi's format
  console.log('Full webhook data:', JSON.stringify(req.body, null, 2));

  // Parse message and phone number from Maytapi webhook
  const message = req.body.message?.text || req.body.message?.body;
  const from = req.body.user?.phone;

  console.log('Webhook received:', message, 'from', from);

  // Initialize user state if not exists
  if (!userStates[from]) {
    userStates[from] = { currentMenu: 'main' };
  }

  // Handle main menu trigger
  if (message && message.trim() === '/') {
    userStates[from].currentMenu = 'main';
    const mainMenu = `🏠 *MAIN MENU*
Please select an option:

1️⃣ Ticket
2️⃣ Order Query  
3️⃣ Stock Query
4️⃣ Document

_Type the number to continue..._`;
    
    await sendWhatsAppMessage(from, mainMenu);
    return res.sendStatus(200);
  }

  // Handle main menu selections
  if (userStates[from] && userStates[from].currentMenu === 'main') {
    if (message && message.trim() === '1') {
      userStates[from].currentMenu = 'ticket';
      const ticketSubMenu = `🎫 *TICKET MENU*
Choose your option:

1️⃣ Help Ticket
2️⃣ Leave Form  
3️⃣ Delegation

_Type the number or click the links below:_`;
      
      await sendWhatsAppMessage(from, ticketSubMenu);
      return res.sendStatus(200);
    }

    if (message && message.trim() === '2') {
      await sendWhatsAppMessage(from, '🔍 *ORDER QUERY*\nThis feature is coming soon!\n\nType */* to return to main menu.');
      return res.sendStatus(200);
    }

    if (message && message.trim() === '3') {
      await sendWhatsAppMessage(from, '📊 *STOCK QUERY*\nThis feature is coming soon!\n\nType */* to return to main menu.');
      return res.sendStatus(200);
    }

    if (message && message.trim() === '4') {
      await sendWhatsAppMessage(from, '📄 *DOCUMENT*\nThis feature is coming soon!\n\nType */* to return to main menu.');
      return res.sendStatus(200);
    }
  }

  // Handle ticket submenu selections
  if (userStates[from] && userStates[from].currentMenu === 'ticket') {
    if (message && message.trim() === '1') {
      const helpTicketMsg = `🆘 *HELP TICKET*

Click the link below to access Help Ticket:
${links.helpTicket}

Type */* to return to main menu.`;
      
      await sendWhatsAppMessage(from, helpTicketMsg);
      userStates[from].currentMenu = 'main'; // Reset to main menu
      return res.sendStatus(200);
    }

    if (message && message.trim() === '2') {
      const leaveFormMsg = `🏖️ *LEAVE FORM*

Click the link below to access Leave Form:
${links.leave}

Type */* to return to main menu.`;
      
      await sendWhatsAppMessage(from, leaveFormMsg);
      userStates[from].currentMenu = 'main'; // Reset to main menu
      return res.sendStatus(200);
    }

    if (message && message.trim() === '3') {
      const delegationMsg = `👥 *DELEGATION*

Click the link below to access Delegation:
${links.delegation}

Type */* to return to main menu.`;
      
      await sendWhatsAppMessage(from, delegationMsg);
      userStates[from].currentMenu = 'main'; // Reset to main menu
      return res.sendStatus(200);
    }
  }

  // Handle invalid input
  if (message && !['/', '1', '2', '3', '4'].includes(message.trim())) {
    await sendWhatsAppMessage(from, '❌ Invalid option. Type */* to see the main menu.');
  }

  res.sendStatus(200);
});

async function sendWhatsAppMessage(to, message) {
  // Debug environment variables
  console.log('Environment Variables Check:');
  console.log('MAYTAPI_PRODUCT_ID:', process.env.MAYTAPI_PRODUCT_ID || 'UNDEFINED');
  console.log('MAYTAPI_PHONE_ID:', process.env.MAYTAPI_PHONE_ID || 'UNDEFINED');
  console.log('MAYTAPI_API_TOKEN:', process.env.MAYTAPI_API_TOKEN ? 'SET (length: ' + process.env.MAYTAPI_API_TOKEN.length + ')' : 'UNDEFINED');

  try {
    console.log('Sending API request with:');
    console.log('Product ID:', process.env.MAYTAPI_PRODUCT_ID);
    console.log('Phone ID:', process.env.MAYTAPI_PHONE_ID);
    console.log('To:', to);
    console.log('Message:', message);

    const response = await axios.post(
      `https://api.maytapi.com/api/${process.env.MAYTAPI_PRODUCT_ID}/${process.env.MAYTAPI_PHONE_ID}/sendMessage`,
      {
        to_number: to,
        type: "text",
        message: message
      },
      {
        headers: {
          'x-maytapi-key': process.env.MAYTAPI_API_TOKEN,
          'Content-Type': 'application/json'
        }
      }
    );
    console.log('Message sent successfully:', response.data);
  } catch (error) {
    console.error('Primary API call failed:', error.response?.data || error.message);
    
    // Try alternative header format
    try {
      console.log('Trying alternative header format...');
      const altResponse = await axios.post(
        `https://api.maytapi.com/api/${process.env.MAYTAPI_PRODUCT_ID}/${process.env.MAYTAPI_PHONE_ID}/sendMessage`,
        {
          to_number: to,
          type: "text",
          message: message
        },
        {
          headers: {
            'X-Maytapi-Key': process.env.MAYTAPI_API_TOKEN, // Capitalized header
            'Content-Type': 'application/json'
          }
        }
      );
      console.log('Alternative header format success:', altResponse.data);
    } catch (altError) {
      console.error('Alternative header format also failed:', altError.response?.data || altError.message);
      
      // Try with different payload format
      try {
        console.log('Trying different payload format...');
        const finalResponse = await axios.post(
          `https://api.maytapi.com/api/${process.env.MAYTAPI_PRODUCT_ID}/${process.env.MAYTAPI_PHONE_ID}/sendMessage`,
          {
            to: to,
            message: message,
            type: "text"
          },
          {
            headers: {
              'x-maytapi-key': process.env.MAYTAPI_API_TOKEN,
              'Content-Type': 'application/json'
            }
          }
        );
        console.log('Different payload format success:', finalResponse.data);
      } catch (finalError) {
        console.error('All API attempts failed:', finalError.response?.data || finalError.message);
      }
    }
  }
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🤖 WhatsApp Bot running on port ${PORT}`);
  console.log('=== ENVIRONMENT VARIABLES DEBUG ===');
  console.log('MAYTAPI_PRODUCT_ID:', process.env.MAYTAPI_PRODUCT_ID || '❌ NOT SET');
  console.log('MAYTAPI_PHONE_ID:', process.env.MAYTAPI_PHONE_ID || '❌ NOT SET');
  console.log('MAYTAPI_API_TOKEN:', process.env.MAYTAPI_API_TOKEN ? '✅ SET' : '❌ NOT SET');
  console.log('NODE_ENV:', process.env.NODE_ENV || 'not set');
  console.log('===================================');
});
