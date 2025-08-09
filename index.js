require('dotenv').config();
const express = require('express');
const axios = require('axios');
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

  // IMMEDIATELY EXIT if message is empty, null, or invalid
  if (!message || typeof message !== 'string') {
    console.log('No valid message - staying silent');
    return res.sendStatus(200);
  }

  const trimmedMessage = message.trim();

  // IMMEDIATELY EXIT if message is empty after trimming
  if (trimmedMessage === '') {
    console.log('Empty message after trim - staying silent');
    return res.sendStatus(200);
  }

  // ONLY process messages that are EXACTLY "/" or menu options from users already in menu
  
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
      userStates[from].currentMenu = 'completed'; // Mark as completed
      return res.sendStatus(200);
    }

    if (trimmedMessage === '2') {
      await sendWhatsAppMessage(from, '🔍 *ORDER QUERY*\nThis feature is coming soon!\n\nType */* to return to main menu.', productId, phoneId);
      userStates[from].currentMenu = 'completed'; // Mark as completed
      return res.sendStatus(200);
    }

    if (trimmedMessage === '3') {
      await sendWhatsAppMessage(from, '📊 *STOCK QUERY*\nThis feature is coming soon!\n\nType */* to return to main menu.', productId, phoneId);
      userStates[from].currentMenu = 'completed'; // Mark as completed
      return res.sendStatus(200);
    }

    if (trimmedMessage === '4') {
      await sendWhatsAppMessage(from, '📄 *DOCUMENT*\nThis feature is coming soon!\n\nType */* to return to main menu.', productId, phoneId);
      userStates[from].currentMenu = 'completed'; // Mark as completed
      return res.sendStatus(200);
    }

    // Invalid menu option - but only for users already in menu
    await sendWhatsAppMessage(from, '❌ Invalid option. Please select 1, 2, 3, or 4.\n\nType */* to see the main menu again.', productId, phoneId);
    return res.sendStatus(200);
  }

  // FOR ALL OTHER MESSAGES: COMPLETE SILENCE
  console.log('Normal message received - bot staying silent:', trimmedMessage);
  return res.sendStatus(200);
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
  console.log('✅ Bot will ONLY respond to "/" commands - complete silence for all other messages');
});
