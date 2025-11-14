📋 Overview
This bot connects Google Sheets with Telegram to create an interactive product/service catalog with search, navigation, and notification features.

🚀 Setup Instructions
Step 1: Google Sheets Preparation
Create a Google Spreadsheet with these required sheets:

Settings - Configuration values

Users - User management

Pokes - Notification system (auto-created)

Your data sheets (categories, products, etc.)

Settings Sheet Structure (Columns A-B):

text
A1: Key              B1: Value
A2: TELEGRAM_TOKEN   B2: YOUR_BOT_TOKEN_HERE
A3: WELCOME_TEXT     B3: Welcome message...
A4: Enable_Password  B4: yes/no
A5: Password         B5: your_password
A6: Poke_Chat_ID     B6: ADMIN_CHAT_ID
... (other config keys)
Step 2: Bot Token Configuration
Where to put the bot token:

In your Google Sheet, go to the Settings tab

Find the row with TELEGRAM_TOKEN in column A

Paste your bot token in column B of the same row

How to get bot token:

Message @BotFather on Telegram

Send /newbot command

Follow instructions to create your bot

Copy the token provided by BotFather

Step 3: API Deployment
Open Google Apps Script:

In your Google Sheet: Extensions > Apps Script

Replace the default code with the provided script

Deploy as Web App:

Click Deploy > New deployment

Type: Web app

Execute as: Me

Who has access: Anyone

Click Deploy

Copy the Web App URL - this is your webhook URL

Set Telegram Webhook:

Open this URL in browser (replace tokens):

text
https://api.telegram.org/bot<YOUR_BOT_TOKEN>/setWebhook?url=<YOUR_WEB_APP_URL>
You should see {"ok":true,"result":true}

Step 4: Required Configuration Keys
Essential Settings (must be in Settings sheet):

TELEGRAM_TOKEN - Your bot token from BotFather

WELCOME_TEXT - Welcome message for /start

Enable_Password - "yes" or "no"

Password - If password enabled

TEXT_SEARCH_PROMPT - Text for search prompt

Poke_Chat_ID - Admin chat ID for notifications (optional)

Step 5: Data Structure
Main Categories:

Sheets with 🔹 in name are treated as categories

Column A: Subcategory sheet names (must contain 🔸)

Column B: Display names

Data Sheets:

Column A: Item names

Other columns: Item details

First row: Headers

🎯 Usage Commands
/start - Begin interaction

/categories - Show main menu

/search - Search items

/help - Show help

⚙️ Key Features
🔍 Search System
Global search across all sheets

Paginated results

Real-time filtering

📱 Navigation
Hierarchical menus (Categories → Subcategories → Items)

Pagination for large datasets

Back navigation

🔔 Poke System
Users can "poke" for item notifications

Admins receive alerts in specified chat

Reply system for admin-user communication

🔐 Security
User authorization system

Password protection option

Blocklist functionality

🛠️ Maintenance
Cache Management
User authorization cached for 5 minutes

Manual cache reset available

Error Handling
Comprehensive error messages

Timeout protection for large datasets

Graceful failure recovery

❗ Important Notes
Sheet Names:

Use 🔹 for category sheets

Use 🔸 for subcategory sheets

Avoid special characters in regular sheet names

Data Limits:

Telegram callback data: 64 characters max

Message length: 4096 characters

Processing timeout: 25 seconds

Required Permissions:

Google Sheets: read/write

UrlFetch: for Telegram API calls

Properties Service: for temporary storage

🆘 Troubleshooting
Common Issues:

"Settings sheet not found" - Create Settings sheet

Webhook errors - Check token and deployment URL

Permission denied - Review Google Apps Script permissions

Timeout errors - Reduce dataset size or optimize sheets

Logs:

Check Apps Script logs: View > Logs

Monitor execution transcripts

