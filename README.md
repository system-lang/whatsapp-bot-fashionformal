# WhatsApp Bot for Fashion Formal

## Features
- Hierarchical menu system (Main menu → Sub menus)
- Clickable links for forms
- User state management
- Error handling

## Menu Structure
/ → Main Menu
├── 1 → Ticket Menu
│ ├── 1 → Help Ticket Link
│ ├── 2 → Leave Form Link
│ └── 3 → Delegation Link
├── 2 → Order Query (Coming Soon)
├── 3 → Stock Query (Coming Soon)
└── 4 → Document (Coming Soon)

text

## Setup Instructions

### 1. Environment Variables
Copy `.env.example` to `.env` and fill in your Maytapi credentials:
MAYTAPI_PRODUCT_ID=your_actual_product_id
MAYTAPI_PHONE_ID=your_actual_phone_id
MAYTAPI_API_TOKEN=your_actual_api_token

text

### 2. Deployment
This bot is configured for Railway deployment.

## Bot Usage
- Send `/` to start
- Use numbers (1, 2, 3, 4) to navigate
- Click links to open forms
- Type `/` anytime to return to main menu

## Technical Stack
- **Runtime**: Node.js
- **Framework**: Express.js
- **HTTP Client**: Axios
- **Hosting**: Railway
- **WhatsApp API**: Maytapi

## Environment Variables Required

| Variable | Description |
|----------|-------------|
| `MAYTAPI_PRODUCT_ID` | Your Maytapi Product ID |
| `MAYTAPI_PHONE_ID` | Your Maytapi Phone Instance ID |
| `MAYTAPI_API_TOKEN` | Your Maytapi API Token |

---
*WhatsApp Bot for Fashion Formal*
