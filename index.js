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
  leave: 'YOUR_LEAVE_FORM_LINK_HERE'
};

// Your API Token
const MAYTAPI_API_TOKEN = '07d75e68-b94f-485b-9e8c-19e707d176ae';

app.post('/webhook', async (req, res) => {
  console.log('Full webhook data:', JSON.stringify(req.body, null, 2));

  // Extract everything we need FROM THE WEBHOOK ITSELF!
  const message = req.body.message?.text;
  const from = req.body.user?.phone;
  const productId = req.body.product_id || req.body.productId;
  const phoneId = req.body.phone_id || req.body.phoneId;

  console.log('Extracted from webhook:');
  console.log('Message:', message);
  console.log('From:', from);
  console.log('Product ID:', productId);
  console.log('Phone ID:', phoneId);

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
    
    await sendWhatsAppMessage(from, mainMenu, productId, phoneId);
    return res.sendStatus(200);
  }

  // Handle main menu selections
  if (userStates[from] && userStates[from].currentMenu === 'main') {
    if (message && message.trim() === '1') {
      // IMPROVED: Direct clickable links instead of submenu
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
      return res.sendStatus(200);
    }

    if (message && message.trim() === '2') {
      await sendWhatsAppMessage(from, '🔍 *ORDER QUERY*\nThis feature is coming soon!\n\nType */* to return to main menu.', productId, phoneId);
      return res.sendStatus(200);
    }

    if (message && message.trim() === '3') {
      await sendWhatsAppMessage(from, '📊 *STOCK QUERY*\nThis feature is coming soon!\n\nType */* to return to main menu.', productId, phoneId);
      return res.sendStatus(200);
    }

    if (message && message.trim() === '4') {
      await sendWhatsAppMessage(from, '📄 *DOCUMENT*\nThis feature is coming soon!\n\nType */* to return to main menu.', productId, phoneId);
      return res.sendStatus(200);
    }
  }

  // Handle invalid input
  if (message && !['/', '1', '2', '3', '4'].includes(message.trim())) {
    await sendWhatsAppMessage(from, '❌ Invalid option. Type */* to see the main menu.');
  }

  res.sendStatus(200);
});

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
  console.log('✅ Using webhook data for API credentials - No environment variables needed!');
});
