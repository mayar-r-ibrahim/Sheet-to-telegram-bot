const MAX_BUTTONS_PER_PAGE = 8;
const CALLBACK_PREFIX = {
  SHEETS: 'sht_',
  LINK_SHEETS: 'lnk_', // For linked spreadsheet sheets
  LINK_NAMES: 'lkn_',  // For names in linked sheets
  NAMES: 'nme_',
  BACK: 'bck_',
  PAGE: 'pg_',
  POKE: 'pok_',
  POKE_REPLY: 'rep_',
  POKE_RESOLVE: 'rsv_'


};

/* ------------------------------------------------------------------------ CACHE FUNCTION  */
// Cache expiration time (5 minutes)
const CACHE_EXPIRATION = 300;

/* ------------------------------------------------------------------------ UTILITY FUNCTIONS  */
// Truncates callback data to Telegram's 64 character limit
function safeCallbackData(data) {
  if (data.length > 64) return data.substring(0, 64);
  return data;
}
// Escapes special Markdown characters for Telegram messages
function escapeMarkdown(text) {
  if (typeof text !== 'string') return text;
  return text.replace(/[_*[\]()~`>#+\-={}.!]/g, '\\$&');
}

/* ------------------------------------------------------------------------ CORE FUNCTIONS  */
// Checks if user is authorized with caching mechanism
function isUserAuthorized(chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `auth_${chatId}`;
  const cached = cache.get(cacheKey);
  
  if (cached !== null) return cached === 'true';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const settingsSheet = ss.getSheetByName('Settings');
  let authorized = true;

  // Check blocked users
  if (usersSheet) {
   const blocked = usersSheet.getRange('O2:O').getValues().flat().map(String).filter(Boolean);
    if (blocked.includes(String(chatId))) {
      authorized = false;
    }
  }

  // Check dropship setting
  if (authorized && settingsSheet) {
    const dropshipEnabled = settingsSheet.getRange('B33').getValue() === 'yes';
    if (dropshipEnabled) {
      const allowed = usersSheet ? usersSheet.getRange('M2:M').getValues().flat().map(String).filter(Boolean) : [];
      authorized = allowed.includes(String(chatId));
    }
  }

  cache.put(cacheKey, authorized.toString(), CACHE_EXPIRATION);
  return authorized;
}

// Retrieves and validates configuration from Settings sheet
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) throw new Error('Settings sheet not found');

  const config = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2)
    .getValues()
    .reduce((acc, [key, value]) => {
      acc[key] = value;
      return acc;
    }, {});

  // Validate required keys
  const requiredKeys = ['Enable_Password', 'Password_Prompt', 'WELCOME_TEXT', 'TELEGRAM_TOKEN', 'TEXT_SEARCH_PROMPT'];
  requiredKeys.forEach(key => {
    if (!config[key]) throw new Error(`Missing required key in Settings: ${key}`);
  });

  return config;
}

// Main entry point for Telegram webhook requests
function doPost(e) {
  try {
    const CONFIG = getConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contents = JSON.parse(e.postData.contents);
    
    if (contents.message) handleMessage(contents.message, ss, CONFIG);
    else if (contents.callback_query) handleCallbackQuery(contents.callback_query, ss, CONFIG);
  } catch (error) {
    Logger.log('Error: ' + error.message);
  }
}

// Handles incoming messages and command routing
function handleMessage(message, ss, CONFIG) {
    const text = message.text || '';
    const chatId = message.chat.id;
  
    // Store user ID
    storeUserId(chatId);
  
    // Authorization check
    if (!isUserAuthorized(chatId)) {
      sendErrorMessage(chatId, CONFIG.Un_Authorized_Text, CONFIG);
      return;
    }
  
    if (/^\/start$/i.test(text)) {
      if (CONFIG.Enable_Password === 'yes') {
        sendTelegramMessage({
          method: 'sendMessage',
          chat_id: chatId,
          text: CONFIG.Password_Prompt,
          reply_markup: { force_reply: true }
        }, CONFIG);
      } else {
        sendWelcomeMessage(chatId, CONFIG);
      }
      return;
    }
  
    if (/^\/categories$/i.test(text)) {
      sendMainMenu(chatId, ss, CONFIG);
      return;
    }
  
    if (message.reply_to_message?.text === (CONFIG.TEXT_POKE_REPLY_PROMPT || '📝 Enter your reply to the user:')) {
      handlePokeReplyMessage(message, CONFIG);
      return;
    }
    
    if (CONFIG.Enable_Password === 'yes' && message.reply_to_message?.text === CONFIG.Password_Prompt) {
      if (text === CONFIG.Password) {
      sendWelcomeMessage(chatId, CONFIG);
      } else {
      sendErrorMessage(chatId, CONFIG.Un_Authorized_Text, CONFIG);
      }
      return;
    }

    if (/^\/search:/.test(text)) {
      const searchQuery = text.replace(/^\/search:/, '').trim();
      handleSearchCommand(chatId, ss, searchQuery, CONFIG);
      return;
    }

    if (message.reply_to_message?.text === CONFIG.TEXT_POKE_REPLY_PROMPT) {
      handlePokeReplyMessage(message, CONFIG);
      return;
    }

    if (/^\/search$/i.test(text)) {
        sendTelegramMessage({
            method: 'sendMessage',
            chat_id: chatId,
            text: CONFIG.TEXT_SEARCH_PROMPT,
            reply_markup: { force_reply: true }
        }, CONFIG);
      return;
    }
    // Handle search input from the search prompt
    if (message.reply_to_message?.text === CONFIG.TEXT_SEARCH_PROMPT) {
      handleSearchCommand(chatId, ss, text, CONFIG);
      return;
    }

  handleRegularCommands(text, chatId, ss, CONFIG);
}

// Handles callback queries from inline keyboards
function handleCallbackQuery(callbackQuery, ss, CONFIG) {
  const data = callbackQuery.data;
  const chatId = callbackQuery.message.chat.id;
  const messageId = callbackQuery.message.message_id;
  
  try {
    if (data === 'main_menu') {
      sendMainMenu(chatId, ss, CONFIG);
    } else if (data.startsWith(CALLBACK_PREFIX.SHEETS)) {
      handleSheetSelection(data, chatId, messageId, ss, CONFIG);
    } else if (data.startsWith(CALLBACK_PREFIX.LINK_SHEETS)) {
      handleLinkedSheetSelection(data, chatId, messageId, CONFIG);
    } else if (data.startsWith(CALLBACK_PREFIX.LINK_NAMES)) {
      handleLinkedNameSelection(data, chatId, messageId, CONFIG);
    } else if (data.startsWith(CALLBACK_PREFIX.NAMES)) {
      handleNameSelection(data, chatId, messageId, ss, CONFIG);
    } else if (data.startsWith(CALLBACK_PREFIX.BACK)) {
      handleBackNavigation(data, chatId, ss, CONFIG);
    } else if (data.startsWith(CALLBACK_PREFIX.PAGE)) {
      handlePagination(data, chatId, messageId, ss, CONFIG);
    } else if (data === 'search_prompt') {
      sendTelegramMessage({
        method: 'sendMessage',
        chat_id: chatId,
        text: CONFIG.TEXT_SEARCH_PROMPT,
        reply_markup: { force_reply: true }
      }, CONFIG);
    } else if (data.startsWith(CALLBACK_PREFIX.POKE)) {
      try {
        handlePokeCommand(data, callbackQuery, CONFIG);
      } catch (pokeError) {
        Logger.log('Poke command failed: ' + pokeError.message);
        sendErrorMessage(chatId, CONFIG.ERROR_POKE_FAILED || 'Failed to send poke notification', CONFIG);
      }
    } else if (data.startsWith(CALLBACK_PREFIX.POKE_REPLY)) {
      handlePokeReply(data.replace(CALLBACK_PREFIX.POKE_REPLY, ''), callbackQuery, CONFIG);
    } else if (data.startsWith(CALLBACK_PREFIX.POKE_RESOLVE)) {
      handlePokeResolve(data.replace(CALLBACK_PREFIX.POKE_RESOLVE, ''), callbackQuery, CONFIG);
    }

  } catch (error) {
    sendErrorMessage(chatId, `Operation failed: ${error.message}`, CONFIG);
  }
}

/* ------------------------------------------------------------------------ POKE SYSTEM  */
// Initializes Pokes sheet structure
function initializePokesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Pokes');



  if (!sheet) {
    sheet = ss.insertSheet('Pokes');
    sheet.appendRow([
      'UUID', // NEW COLUMN
      'Timestamp', 
      'User ID', 
      'Customer', 
      'Sheet', 
      'Linked Spreadsheet ID',
      'Status',
      'Admin ID',
      'Admin Reply',
      'Reply Timestamp'
    ]);
    sheet.hideSheet();
  }
  // Add UUID column if missing
  else if (sheet.getRange(1, 1).getValue() !== 'UUID') {
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue('UUID');
  }
}

// Handles poke resolution by admins
function handlePokeResolve(data, callbackQuery, CONFIG) {
  const uuid = data.replace(CALLBACK_PREFIX.POKE_RESOLVE, '');
  const pokeData = getPokeDetails(uuid);
  
  if (!pokeData) {
    sendErrorMessage(callbackQuery.message.chat.id, 'Poke not found', CONFIG);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pokeSheet = ss.getSheetByName('Pokes');
  const dataRange = pokeSheet.getDataRange().getValues();
  const header = dataRange[0];
  
  for (let i = dataRange.length - 1; i >= 0; i--) {
    if (dataRange[i][header.indexOf('UUID')] === uuid) {
      const row = i + 1;
      pokeSheet.getRange(row, header.indexOf('Status') + 1).setValue('Resolved');
      pokeSheet.getRange(row, header.indexOf('Admin ID') + 1).setValue(callbackQuery.from.id);
      pokeSheet.getRange(row, header.indexOf('Reply Timestamp') + 1).setValue(new Date());
      break;
    }
  }


    // Edit the original message to remove buttons and show resolved status
    const originalText = callbackQuery.message.text;
    const resolvedText = originalText + '\n\n✅ *Resolved*';
  

  // Confirm resolution
  sendTelegramMessage({
    method: 'editMessageText',
    chat_id: callbackQuery.message.chat.id,
    message_id: callbackQuery.message.message_id,
    text: resolvedText,
    parse_mode: 'Markdown',
    reply_markup: { inline_keyboard: [] } // removes all buttons
  }, CONFIG);
}

// Processes poke commands from users
function handlePokeCommand(data, callbackQuery, CONFIG) {
  const pokeData = data.replace(CALLBACK_PREFIX.POKE, '');
  const parts = pokeData.split('|');
  const type = parts[0];
  let sheetName, itemName, linkedSpreadsheetId;

  // Validate and parse poke type
  if (type === 'local') {
    [sheetName, itemName] = parts.slice(1, 3);
  } else if (type === 'linked') {
    [linkedSpreadsheetId, sheetName, itemName] = parts.slice(1, 4);
  } else {
    Logger.log(`Invalid poke type: ${type}`);
    return;
  }

  // Generate UUID and prepare data
  const uuid = Utilities.getUuid();
  const user = callbackQuery.from;
  const timestamp = new Date();

  // Store in Pokes sheet (single append operation)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pokeSheet = ss.getSheetByName('Pokes') || initializePokesSheet();
  pokeSheet.appendRow([
    uuid,                     // UUID
    timestamp,                // Timestamp
    user.id,                  // User ID
    itemName,                 // Customer
    sheetName,                // Sheet
    linkedSpreadsheetId || 'N/A', // Linked Spreadsheet ID
    'Pending',                // Status
    '',                       // Admin ID
    '',                       // Admin Reply
    ''                        // Reply Timestamp
  ]);

  // Prepare admin message
  const userName = user.username 
    ? `@${user.username}` 
    : `${user.first_name} (ID: ${user.id})`;
  
  const messageText = '🔔 *Poke Alert\!*\n' +
    `*User:* ${escapeMarkdown(userName)}\n` +
    `*Customer:* ${escapeMarkdown(itemName)}\n` +
    `*From:* ${escapeMarkdown(sheetName)}` +
    (type === 'linked' ? `\n*Linked Spreadsheet ID:* \`${linkedSpreadsheetId}\`` : '');

  // Create callback buttons
  const replyData = `${CALLBACK_PREFIX.POKE_REPLY}${uuid}`;
  const resolveData = `${CALLBACK_PREFIX.POKE_RESOLVE}${uuid}`;

  // Send admin alert
  if (CONFIG.Poke_Chat_ID) {
    sendTelegramMessage({
      method: 'sendMessage',
      chat_id: CONFIG.Poke_Chat_ID,
      text: messageText,
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [[
          { text: '✉️ Reply', callback_data: replyData },
          { text: '✅ Resolve', callback_data: resolveData }
        ]]
      }
    }, CONFIG);
  } else {
    Logger.log('Poke_Chat_ID not configured in Settings.');
  }

  // Send user confirmation
  const confirmationText = CONFIG.TEXT_POKE_SENT || 'Your poke has been sent.';
  sendTelegramMessage({
    method: 'sendMessage',
    chat_id: callbackQuery.message.chat.id,
    text: confirmationText
  }, CONFIG);
}
// Handles admin replies to pokes
function handlePokeReplyMessage(message, CONFIG) {
  const text = message.text;
  const adminId = message.from.id;
  const scriptProperties = PropertiesService.getScriptProperties();
  const context = scriptProperties.getProperty(`POKE_${adminId}`);
  const MAX_REPLY_LENGTH = 4000;
  const truncatedText = text.length > MAX_REPLY_LENGTH 
  ? text.substring(0, MAX_REPLY_LENGTH) + '...' 
  : text;


  if (!context) {
    Logger.log('No poke context found for admin:', adminId);
    return;
  }

  // Parse and immediately delete the stored context
  const { uuid } = JSON.parse(context);
  scriptProperties.deleteProperty(`POKE_${adminId}`);

  // Get fresh data from sheet
  const pokeData = getPokeDetails(uuid);
  if (!pokeData) {
    Logger.log('Poke data not found for UUID:', uuid);
    sendErrorMessage(adminId, CONFIG.ERROR_POKE_NOT_FOUND, CONFIG);
    return;
  }

  // Send reply to user
  const replyText = `📩 *New Reply from Admin*\n` +
                   `Regarding: ${escapeMarkdown(pokeData.itemName)} ` +
                   `(${escapeMarkdown(pokeData.sheetName)})\n\n` +
                   escapeMarkdown(text);

  sendTelegramMessage({
    method: 'sendMessage',
    chat_id: pokeData.userId,
    text: replyText,
    parse_mode: 'Markdown'
  }, CONFIG);

  // Update Pokes sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pokeSheet = ss.getSheetByName('Pokes');
  const [header, ...rows] = pokeSheet.getDataRange().getValues();
  
  const statusCol = header.indexOf('Status') + 1;
  const adminIdCol = header.indexOf('Admin ID') + 1;
  const adminReplyCol = header.indexOf('Admin Reply') + 1;
  const replyTimeCol = header.indexOf('Reply Timestamp') + 1;

  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][header.indexOf('UUID')] === uuid) {
      const row = i + 2; // +2 because header is row 1 (1-based) and rows start from index 0
      pokeSheet.getRange(row, statusCol).setValue('Replied');
      pokeSheet.getRange(row, adminIdCol).setValue(adminId);
      pokeSheet.getRange(row, adminReplyCol).setValue(text);
      pokeSheet.getRange(row, replyTimeCol).setValue(new Date());
      break;
    }
  }

  // Confirm to admin
  sendTelegramMessage({
    method: 'sendMessage',
    chat_id: adminId,
    text: CONFIG.TEXT_POKE_REPLY_SENT || '✅ Reply sent successfully'
  }, CONFIG);
}

function handlePokeReply(data, callbackQuery, CONFIG) {
  const uuid = data.replace(CALLBACK_PREFIX.POKE_REPLY, '');
  const adminId = callbackQuery.from.id;

  // Get details from Pokes sheet
  const pokeData = getPokeDetails(uuid);
  if (!pokeData) {
    sendErrorMessage(callbackQuery.message.chat.id, '❌ Poke request not found or expired', CONFIG);
    return;
  }

  // Store context in PropertiesService
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(`POKE_${adminId}`, JSON.stringify({
    uuid: uuid,
    timestamp: new Date().toISOString(),
    targetUser: pokeData.userId,
    itemName: pokeData.itemName,
    sheetName: pokeData.sheetName,
    linkedId: pokeData.linkedId
  }));

  // Prepare and send reply prompt with proper force reply
  const promptText = CONFIG.TEXT_POKE_REPLY_PROMPT || '📝 Enter your reply to the user:';
  sendTelegramMessage({
    method: 'sendMessage',
    chat_id: callbackQuery.message.chat.id,
    text: promptText,
    reply_markup: {
      force_reply: true,
      selective: true,
      input_field_placeholder: 'Type your response here...'
    }
  }, CONFIG);
}


//  helper function to get poke details
function getPokeDetails(uuid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pokeSheet = ss.getSheetByName('Pokes');
  
  if (!pokeSheet) {
    Logger.log('Pokes sheet not found');
    return null;
  }

  const data = pokeSheet.getDataRange().getValues();
  if (data.length < 2) return null; // No records
  
  const header = data[0];
  // Calculate column indices AFTER header is defined
  const UUID_COL = header.indexOf('UUID');
  const USER_ID_COL = header.indexOf('User ID');
  const CUSTOMER_COL = header.indexOf('Customer');
  const SHEET_COL = header.indexOf('Sheet');
  const LINKED_ID_COL = header.indexOf('Linked Spreadsheet ID');

  // Find row with matching UUID using pre-calculated columns
  for (let i = 1; i < data.length; i++) {
    if (data[i][UUID_COL] === uuid) {
      return {
        userId: data[i][USER_ID_COL],
        itemName: data[i][CUSTOMER_COL],
        sheetName: data[i][SHEET_COL],
        linkedId: data[i][LINKED_ID_COL]
      };
    }
  }
  return null;
}





/* ------------------------------------------------------------------------ SEARCH SYSTEM  */
// Handles search command execution and results display
function handleSearchCommand(chatId, ss, searchQuery, CONFIG, page = 0, messageId = null) {
  const startTime = new Date().getTime();
  
  if (!searchQuery || searchQuery.trim() === '') {
    sendErrorMessage(chatId, CONFIG.TEXT_NO_SEARCH_QUERY, CONFIG);
    return;
  }

  const sheets = ss.getSheets().filter(s => !['Settings', 'Users', 'Statistics'].includes(s.getName()));
  const results = [];
  const maxProcessingTime = 25000; // 25 sec

  // Process sheets in batches
  for (const sheet of sheets) {
    // Check timeout
    if (new Date().getTime() - startTime > maxProcessingTime) {
      sendErrorMessage(chatId, CONFIG.ERROR_TIMEOUT, CONFIG);
      return;
    }

    try {
      const names = sheet.getRange('A2:A' + sheet.getLastRow()).getValues()
        .flat()
        .filter(name => Boolean(name) && typeof name === 'string');
      
      const matches = names.filter(name => 
        name.toLowerCase().includes(searchQuery.toLowerCase())
      );
      
      matches.forEach(match => {
    results.push({
      sheetName: sheet.getName(), // Store the sheet name string
      name: match,
      callback: safeCallbackData(`${CALLBACK_PREFIX.NAMES}${sheet.getName()}|${match}`)
    });
  });



  const keyboard = paginatedResults.map(result => [{
    text: `${result.name} (${result.sheetName})`, // Use the stored sheet name directly
    callback_data: result.callback
  }]);


    } catch (e) {
      Logger.log(`Error processing sheet ${sheet.getName()}: ${e}`);
      continue;
    }
  }

  const resultsPerPage = MAX_BUTTONS_PER_PAGE;
  const totalPages = results.length > 0 ? Math.ceil(results.length / resultsPerPage) : 0;
  
  if (results.length === 0) {
    const noResultsText = CONFIG.TEXT_NO_RESULTS_FOUND.includes('{query}') 
      ? CONFIG.TEXT_NO_RESULTS_FOUND.replace('{query}', searchQuery)
      : `${CONFIG.TEXT_NO_RESULTS_FOUND}: "${searchQuery}"`;
    
    sendTelegramMessage({
      method: messageId ? 'editMessageText' : 'sendMessage',
      chat_id: chatId,
      message_id: messageId,
      text: noResultsText
    }, CONFIG);
    return;
  }

  page = Math.max(0, Math.min(page, totalPages - 1));
  const paginatedResults = results.slice(page * resultsPerPage, (page + 1) * resultsPerPage);

  const keyboard = paginatedResults.map(result => [{
    text: `${result.name} (${result.sheetName})`, 
    callback_data: result.callback
  }]);

  // Pagination controls
  const pagination = [];
  if (totalPages > 1) {
    if (page > 0) {
      pagination.push({ 
        text: CONFIG.ICON_PREV, 
        callback_data: safeCallbackData(`${CALLBACK_PREFIX.PAGE}search|${page - 1}|${encodeURIComponent(searchQuery)}`)
      });
    }
    if (page < totalPages - 1) {
      pagination.push({ 
        text: CONFIG.ICON_NEXT, 
        callback_data: safeCallbackData(`${CALLBACK_PREFIX.PAGE}search|${page + 1}|${encodeURIComponent(searchQuery)}`)
      });
    }
  }
  
  if (pagination.length) keyboard.push(pagination);
  keyboard.push([{ 
    text: CONFIG.BUTTON_BACK, 
    callback_data: `${CALLBACK_PREFIX.BACK}main`
  }]);

  sendTelegramMessage({
    method: messageId ? 'editMessageText' : 'sendMessage',
    chat_id: chatId,
    message_id: messageId,
    text: `${CONFIG.TEXT_SEARCH_RESULTS} (${page + 1}/${totalPages}):\nSearch: "${searchQuery}"`,
    reply_markup: { inline_keyboard: keyboard }
  }, CONFIG);
}

/* ------------------------------------------------------------------------ MENU SYSTEM  */

//  Main menu display with categories/sheets
function sendMainMenu(chatId, ss, CONFIG, page = 0) {
  const excludedSheets = ['Settings', 'Users', 'Pokes', 'Statistics'];
  const sheets = ss.getSheets().filter(s => 
  !excludedSheets.includes(s.getName()) && 
  !s.getName().includes('🔸')
  );


  const totalPages = Math.ceil(sheets.length / MAX_BUTTONS_PER_PAGE);
  const paginatedSheets = sheets.slice(page * MAX_BUTTONS_PER_PAGE, (page + 1) * MAX_BUTTONS_PER_PAGE);
  
  const keyboard = paginatedSheets.map(sheet => [
  { 
    text: `${sheet.getName().includes('🔹') ? CONFIG.ICON_CATEGORY : CONFIG.ICON_SHEET} ${getDisplayName(sheet)}`, 
    callback_data: safeCallbackData(`${CALLBACK_PREFIX.SHEETS}${sheet.getName()}|0`)
  }
  ]);


  //  Pagination buttons "if needed"
  if (sheets.length > MAX_BUTTONS_PER_PAGE) {
    const pagination = [];
    if (page > 0) pagination.push({ 
      text: CONFIG.ICON_PREV, 
      callback_data: safeCallbackData(`${CALLBACK_PREFIX.PAGE}sheet|${page - 1}`)
    });
    if (page < totalPages - 1) pagination.push({ 
      text: CONFIG.ICON_NEXT, 
      callback_data: `${CALLBACK_PREFIX.PAGE}sheet|${page + 1}`
    });
    keyboard.push(pagination);
  }

  // Search 🔎 button
  keyboard.push([{ text: 'Search 🔎', callback_data: 'search_prompt' }]);
  
  sendOrEditMessage(chatId, {
    text: `${CONFIG.ICON_CATEGORY} ${CONFIG.TEXT_SELECT_CATEGORY} (${CONFIG.TEXT_PAGE} ${page + 1}/${totalPages}):`,
    reply_markup: { inline_keyboard: keyboard }
  }, CONFIG);
}

// Handles sheet/subcategory selection
function handleSheetSelection(data, chatId, messageId, ss, CONFIG) {
  const [sheetName, page] = data.split('|').slice(0, 2);
  const cleanSheetName = sheetName.replace(CALLBACK_PREFIX.SHEETS, '');
  const currentSheet = ss.getSheetByName(cleanSheetName);

  if (!currentSheet) {
      sendErrorMessage(chatId, CONFIG.ERROR_SHEET_NOT_FOUND, CONFIG);
      return;
  }

  if (currentSheet.getName().includes('🔹')) {
      // Handle category sheet (🔹)
      const subcategoryData = currentSheet.getRange('A2:B' + currentSheet.getLastRow()).getValues();
      
      // Validate subcategories with 🔸 check and sheet existence
      const validSubcategories = subcategoryData
          .map(row => {
              const sheetName = String(row[0]).trim();
              const displayName = row[1] ? String(row[1]).trim() : sheetName;
              return { sheetName, displayName };
          })
          .filter(entry => {
              // Check for 🔸 and valid sheet name
              const hasValidFormat = entry.sheetName && entry.sheetName.includes('🔸');
              const sheetExists = ss.getSheetByName(entry.sheetName);
              return hasValidFormat && sheetExists;
          });

      // Map to sheet objects with display names
      const subcategoryEntries = validSubcategories.map(entry => ({
          sheet: ss.getSheetByName(entry.sheetName),
          displayName: entry.displayName
      }));

      if (subcategoryEntries.length === 0) {
          sendErrorMessage(chatId, CONFIG.ERROR_NO_SUBCATEGORIES, CONFIG);
          return;
      }

      sendSubcategoryMenu(chatId, messageId, subcategoryEntries, currentSheet.getName(), CONFIG, 0);
  } else {
      // Handle regular sheet
      sendNameMenu(chatId, messageId, ss, cleanSheetName, parseInt(page), CONFIG);
  }
}

// Subcategory menu Display
function sendSubcategoryMenu(chatId, messageId, subSheets, categoryName, CONFIG, page = 0) {
  const totalPages = Math.ceil(subSheets.length / MAX_BUTTONS_PER_PAGE);
  const paginatedEntries = subSheets.slice(page * MAX_BUTTONS_PER_PAGE, (page + 1) * MAX_BUTTONS_PER_PAGE);

  const keyboard = paginatedEntries.map(entry => [
    { 
      text: `${CONFIG.ICON_SUBCATEGORY} ${entry.displayName}`, 
      callback_data: `${CALLBACK_PREFIX.SHEETS}${entry.sheet.getName()}|0`
    }
  ]);


  // Pagination buttons
  const pagination = [];
  if (subSheets.length > MAX_BUTTONS_PER_PAGE) {
    if (page > 0) pagination.push({
      text: CONFIG.ICON_PREV,
      callback_data: `${CALLBACK_PREFIX.PAGE}subcategory|${page - 1}|${categoryName}`
    });
    if (page < totalPages - 1) pagination.push({
      text: CONFIG.ICON_NEXT,
      callback_data: `${CALLBACK_PREFIX.PAGE}subcategory|${page + 1}|${categoryName}`
    });
  }
  if (pagination.length) keyboard.push(pagination);

  // Back button
  keyboard.push([{ 
    text: CONFIG.BUTTON_BACK, 
    callback_data: `${CALLBACK_PREFIX.BACK}main`
  }]);

  sendOrEditMessage(chatId, {
    message_id: messageId,
    text: `${CONFIG.ICON_CATEGORY} ${categoryName} - ${CONFIG.TEXT_SELECT_SUBCATEGORY} (Page ${page + 1}/${totalPages})`,
    reply_markup: { inline_keyboard: keyboard }
  }, CONFIG);
}

function handleBackNavigation(data, chatId, ss, CONFIG) {
  const backData = data.replace(CALLBACK_PREFIX.BACK, '');
  const [type, sheetName] = backData.split('|');

  switch (type) {
    case 'main':
      sendMainMenu(chatId, ss, CONFIG);
      break;
    case 'names':
      sendNameMenu(chatId, null, ss, sheetName, 0, CONFIG);
      break;
    case 'search':
      const searchQuery = navigationHistory[chatId]?.searchQuery || '';
      handleSearchCommand(chatId, ss, searchQuery, CONFIG);
      break;
    default:
      sendMainMenu(chatId, ss, CONFIG);
  }
}

// Handles navigation back actions
function handleBackNavigation(data, chatId, ss, CONFIG) {
  const backData = data.replace(CALLBACK_PREFIX.BACK, '');
  const [type, ...args] = backData.split('|');
  
  switch (type) {
    case 'main':
      sendMainMenu(chatId, ss, CONFIG);
      break;  
  case 'names':
    sendNameMenu(chatId, null, ss, args[0], 0, CONFIG);
    break;
  case 'linked_sheets':
    const linkedSpreadsheetId = args[0];
    const linkedSS = SpreadsheetApp.openById(linkedSpreadsheetId);
    const linkedSheets = linkedSS.getSheets().filter(s => !['Settings', 'Users', 'Statistics'].includes(s.getName()));
    sendLinkedSheetsMenu(chatId, null, linkedSheets, linkedSpreadsheetId, CONFIG, 0);
    break;
  case 'linked_names':
    const [linkedSSId, sheetName] = args;
    handleLinkedSheetSelection(`${CALLBACK_PREFIX.LINK_SHEETS}${linkedSSId}|${sheetName}|0`, chatId, null, CONFIG);
    break;
  default:
    sendMainMenu(chatId, ss, CONFIG);
}
}

/* ------------------------------------------------------------------------ LINKED SPREADSHEETS  */
// Handles linked spreadsheet sheet selection
function handleLinkedSheetSelection(data, chatId, messageId, CONFIG) {
  const parts = data.replace(CALLBACK_PREFIX.LINK_SHEETS, '').split('|');
  const [linkedSpreadsheetId, sheetName, page] = parts;
  let linkedSS;
  try {
    linkedSS = SpreadsheetApp.openById(linkedSpreadsheetId);
  } catch (e) {
    sendErrorMessage(chatId, CONFIG.ERROR_LINK_INVALID, CONFIG);
    return;
  }
  sendLinkedNameMenu(chatId, messageId, linkedSS, sheetName, parseInt(page), linkedSpreadsheetId, CONFIG);
}

// Displays menu for linked spreadsheet sheets
function sendLinkedNameMenu(chatId, messageId, linkedSS, sheetName, page, linkedSpreadsheetId, CONFIG) {
  const sheet = linkedSS.getSheetByName(sheetName);
  if (!sheet) {
    sendErrorMessage(chatId, CONFIG.ERROR_SHEET_NOT_FOUND, CONFIG);
    return;
  }
  
  const names = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat().filter(Boolean);
  const totalPages = Math.ceil(names.length / MAX_BUTTONS_PER_PAGE);
  const paginatedNames = names.slice(page * MAX_BUTTONS_PER_PAGE, (page + 1) * MAX_BUTTONS_PER_PAGE);
  
  const keyboard = paginatedNames.map(name => [
    { text: `${CONFIG.ICON_NAME} ${name}`, callback_data: `${CALLBACK_PREFIX.LINK_NAMES}${linkedSpreadsheetId}|${sheetName}|${name}` }
  ]);

  const navButtons = [];
  if (names.length > MAX_BUTTONS_PER_PAGE) {
    if (page > 0) navButtons.push({ 
      text: CONFIG.ICON_PREV, 
      callback_data: `${CALLBACK_PREFIX.PAGE}link_name|${page - 1}|${linkedSpreadsheetId}|${sheetName}`
    });
    if (page < totalPages - 1) navButtons.push({ 
      text: CONFIG.ICON_NEXT, 
      callback_data: `${CALLBACK_PREFIX.PAGE}link_name|${page + 1}|${linkedSpreadsheetId}|${sheetName}`
    });
  }
  navButtons.push({ 
    text: CONFIG.BUTTON_BACK, 
    callback_data: `${CALLBACK_PREFIX.BACK}linked_sheets|${linkedSpreadsheetId}`
  });
  
  keyboard.push(navButtons);
  
  sendOrEditMessage(chatId, {
    message_id: messageId,
    text: `${CONFIG.ICON_LIST} ${sheetName} - ${CONFIG.TEXT_SELECT_ITEM} (${CONFIG.TEXT_PAGE} ${page + 1}/${totalPages}):`,
    reply_markup: { inline_keyboard: keyboard }
  }, CONFIG);
}

// Handles displaying linked spreadsheet sheets menu
function sendLinkedSheetsMenu(chatId, messageId, linkedSheets, linkedSpreadsheetId, CONFIG, page = 0) {
  const totalPages = Math.ceil(linkedSheets.length / MAX_BUTTONS_PER_PAGE);
  const paginatedSheets = linkedSheets.slice(page * MAX_BUTTONS_PER_PAGE, (page + 1) * MAX_BUTTONS_PER_PAGE);
  
  const keyboard = paginatedSheets.map(sheet => [
    { 
      text: `${CONFIG.ICON_LINKED} ${getDisplayName(sheet)}`, 
      callback_data: `${CALLBACK_PREFIX.LINK_SHEETS}${linkedSpreadsheetId}|${sheet.getName()}|0`
    }
  ]);


  const pagination = [];
  if (linkedSheets.length > MAX_BUTTONS_PER_PAGE) {
    if (page > 0) pagination.push({ 
      text: CONFIG.ICON_PREV, 
      callback_data: `${CALLBACK_PREFIX.PAGE}link_sheet|${page - 1}|${linkedSpreadsheetId}`
    });
    if (page < totalPages - 1) pagination.push({ 
      text: CONFIG.ICON_NEXT, 
      callback_data: `${CALLBACK_PREFIX.PAGE}link_sheet|${page + 1}|${linkedSpreadsheetId}`
    });
  }
  if (pagination.length) keyboard.push(pagination);
  
  keyboard.push([{ text: CONFIG.BUTTON_BACK, callback_data: `${CALLBACK_PREFIX.BACK}main` }]);
  
  sendOrEditMessage(chatId, {
    message_id: messageId,
    text: `${CONFIG.ICON_LINKED} ${CONFIG.TEXT_SELECT_LINKED_CATEGORY} (${CONFIG.TEXT_PAGE} ${page + 1}/${totalPages}):`,
    reply_markup: { inline_keyboard: keyboard }
  }, CONFIG);
}

// Handles selection of names in linked sheets
function handleLinkedNameSelection(data, chatId, messageId, CONFIG) {
  const parts = data.replace(CALLBACK_PREFIX.LINK_NAMES, '').split('|');
  const [linkedSpreadsheetId, sheetName, name] = parts;
  let linkedSS;
  try {
    linkedSS = SpreadsheetApp.openById(linkedSpreadsheetId);
  } catch (e) {
    sendErrorMessage(chatId, CONFIG.ERROR_LINK_INVALID, CONFIG);
    return;
  }
  
  const sheet = linkedSS.getSheetByName(sheetName);
  if (!sheet) return sendErrorMessage(chatId, CONFIG.ERROR_SHEET_NOT_FOUND, CONFIG);
  
  const dataRange = sheet.getDataRange().getValues();
  const headers = dataRange[0];
  const rowData = dataRange.find(r => r[0] === name);
  
  if (!rowData) return sendErrorMessage(chatId, CONFIG.ERROR_DATA_NOT_FOUND, CONFIG);
  
  // Modified message construction with escaping
const message = headers.reduce((acc, header, index) => {
  const escapedHeader = escapeMarkdown(header);
  const escapedValue = escapeMarkdown(rowData[index] || CONFIG.TEXT_NOT_AVAILABLE);
  return acc + (index === 0 ? '' : `\n*${escapedHeader}:* ${escapedValue}`);
}, `*${escapeMarkdown(name)}*\n`);

  sendOrEditMessage(chatId, {
  message_id: messageId,
  text: message,
  parse_mode: 'Markdown',
  reply_markup: {
    inline_keyboard: [[
      { 
        text: CONFIG.BUTTON_BACK, 
        callback_data: `${CALLBACK_PREFIX.BACK}linked_names|${linkedSpreadsheetId}|${sheetName}`
      },
      { 
        text: CONFIG.BUTTON_POKE, 
        callback_data: safeCallbackData(`${CALLBACK_PREFIX.POKE}linked|${linkedSpreadsheetId}|${sheetName}|${name}`)
      }
    ]]
  }
}, CONFIG);

}

/* ------------------------------------------------------------------------ NAME/ITEM HANDLING  */
// Displays item details for selected name
function handleNameSelection(data, chatId, messageId, ss, CONFIG) {
  const [sheetName, name] = data.split('|').slice(0, 2);
  const cleanSheetName = sheetName.replace(CALLBACK_PREFIX.NAMES, '');
  
  const sheet = ss.getSheetByName(cleanSheetName);
  if (!sheet) return sendErrorMessage(chatId, CONFIG.ERROR_SHEET_NOT_FOUND, CONFIG);
  
  const dataRange = sheet.getDataRange().getValues();
  const headers = dataRange[0];
  const rowData = dataRange.find(r => r[0] === name);
  
  if (!rowData) return sendErrorMessage(chatId, CONFIG.ERROR_DATA_NOT_FOUND, CONFIG);
  
 const message = headers.reduce((acc, header, index) => {
  const escapedHeader = escapeMarkdown(header);
  const escapedValue = escapeMarkdown(rowData[index] || CONFIG.TEXT_NOT_AVAILABLE);
  return acc + (index === 0 ? '' : `\n*${escapedHeader}:* ${escapedValue}`);
}, `*${escapeMarkdown(name)}*\n`);

  
  sendOrEditMessage(chatId, {
  message_id: messageId,
  text: message,
  parse_mode: 'Markdown',
  reply_markup: {
    inline_keyboard: [[
      { text: CONFIG.BUTTON_BACK, callback_data: `${CALLBACK_PREFIX.BACK}names|${cleanSheetName}` },
      { 
        text: CONFIG.BUTTON_POKE, 
        callback_data: safeCallbackData(`${CALLBACK_PREFIX.POKE}local|${cleanSheetName}|${name}`)
      }
    ]]
  }
}, CONFIG);

}

// Shows menu of names/items in a sheet
function sendNameMenu(chatId, messageId, ss, sheetName, page, CONFIG) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return sendErrorMessage(chatId, CONFIG.ERROR_SHEET_NOT_FOUND, CONFIG);
  
  const names = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat().filter(Boolean);
  const totalPages = Math.ceil(names.length / MAX_BUTTONS_PER_PAGE);
  const paginatedNames = names.slice(page * MAX_BUTTONS_PER_PAGE, (page + 1) * MAX_BUTTONS_PER_PAGE);
  
  const keyboard = paginatedNames.map(name => [
    { text: `${CONFIG.ICON_NAME} ${name}`, callback_data: safeCallbackData(`${CALLBACK_PREFIX.NAMES}${sheetName}|${name}`)
  }
  ]);
  
  const navButtons = [];
  if (names.length > MAX_BUTTONS_PER_PAGE) {
    if (page > 0) navButtons.push({ 
      text: CONFIG.ICON_PREV, 
      callback_data: `${CALLBACK_PREFIX.PAGE}name|${page - 1}|${sheetName}`
    });
    if (page < totalPages - 1) navButtons.push({ 
      text: CONFIG.ICON_NEXT, 
      callback_data: `${CALLBACK_PREFIX.PAGE}name|${page + 1}|${sheetName}`
    });
  }
  navButtons.push({ text: CONFIG.BUTTON_BACK, callback_data: `${CALLBACK_PREFIX.BACK}main` });
  
  keyboard.push(navButtons);
  
  sendOrEditMessage(chatId, {
    message_id: messageId,
    text: `${CONFIG.ICON_LIST} ${sheetName} - ${CONFIG.TEXT_SELECT_ITEM} (${CONFIG.TEXT_PAGE} ${page + 1}/${totalPages}):`,
    reply_markup: { inline_keyboard: keyboard }
  }, CONFIG);
}

/* ------------------------------------------------------------------------ MESSAGE HELPERS  */
// Unified message sending/editing function
function sendOrEditMessage(chatId, options, CONFIG) {
  const payload = {
    method: options.message_id ? 'editMessageText' : 'sendMessage',
    chat_id: chatId,
    message_id: options.message_id,
    text: options.text || CONFIG.TEXT_DEFAULT_PROMPT,
    parse_mode: options.parse_mode,
    reply_markup: JSON.stringify(options.reply_markup)
  };
  
  sendTelegramMessage(payload, CONFIG);
}

// Sends error messages to users
function sendErrorMessage(chatId, text, CONFIG) {
  sendTelegramMessage({
    method: 'sendMessage',
    chat_id: chatId,
    text: `${CONFIG.ERROR_PREFIX} ${text}`
  }, CONFIG);
}

// Sends welcome message to new users
function sendWelcomeMessage(chatId, CONFIG) {
  sendOrEditMessage(chatId, {
    text: CONFIG.WELCOME_TEXT,
    parse_mode: 'Markdown',
    reply_markup: {
      inline_keyboard: [[
        { text: CONFIG.BUTTON_VIEW_CATEGORIES, callback_data: 'main_menu' }
      ]]
    }
  }, CONFIG);
}

// Sends help information to users
function sendHelp(chatId, CONFIG) {
  sendTelegramMessage({
    method: 'sendMessage',
    chat_id: chatId,
    text: CONFIG.HELP_TEXT,
    parse_mode: 'Markdown'
  }, CONFIG);
}

/* ------------------------------------------------------------------------ USER MANAGEMENT  */
// Stores user IDs in Users sheet
function storeUserId(chatId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');

  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['Timestamp', 'User ID']);
    sheet.hideSheet();
  }

  const userIds = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
  if (!userIds.includes(chatId)) {
    sheet.appendRow([new Date(), chatId]);
    CacheService.getScriptCache().remove(`auth_${chatId}`); // Invalidate cache
  }
}

// Initializes Users sheet structure
function initializeUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['Timestamp', 'User ID']);
    sheet.hideSheet();
  }
}

/* ------------------------------------------------------------------------ TELEGRAM API  */
// Direct Telegram API communication handler
function sendTelegramMessage(payload, CONFIG) {
  try {
    const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM_TOKEN}/`;
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
    Logger.log(`Telegram API Response: ${response.getContentText()}`);
  } catch (e) {
    Logger.log(`Telegram API Error: ${e.message}. Payload: ${JSON.stringify(payload)}`);
    throw e;
  }
}

/* ------------------------------------------------------------------------ HELPER FUNCTIONS  */
// Generates display names for sheets with special formatting
function getDisplayName(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sheet.getName();

  try {
    // Handle category sheets (🔹)
    if (sheetName.includes('🔹')) {
      const displayName = sheet.getRange('B1').getValue();
      return displayName || sheetName;
    }
    
    // Handle subcategory sheets (🔸)
    if (sheetName.includes('🔸')) {
      const parentSheets = ss.getSheets().filter(s => s.getName().includes('🔹'));
      
      for (const parentSheet of parentSheets) {
        const subcategories = parentSheet.getRange('A2:A' + parentSheet.getLastRow())
          .getValues()
          .flat()
          .map(name => name.toString().trim());
        
        const index = subcategories.indexOf(sheetName);
        if (index !== -1) {
          const displayName = parentSheet.getRange(index + 2, 2).getValue();
          return displayName || sheetName;
        }
      }
    }

  if (!sheet || !sheet.getName) return "Unknown Category"; // safety check
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sheet.getName();

  } catch(e) {
    Logger.log('Error getting display name: ' + e);
  }
  
  return sheetName; // Fallback to original sheet title if no name found
}

// Handles pagination for various menus
function handlePagination(data, chatId, messageId, ss, CONFIG) {
  try {
    const parts = data.split('|');
    const type = parts[0].replace(CALLBACK_PREFIX.PAGE, '');
    const page = parseInt(parts[1]) || 0;
    const query = decodeURIComponent(parts.slice(2).join('|'));

    if (page < 0) return; // Prevent negative pages

    switch (type) {
      case 'search':
        handleSearchCommand(chatId, ss, query, CONFIG, page);
        break;

        case 'subcategory':
          // Get full subcategory data with display names
          const categorySheet = ss.getSheetByName(query);
          if (!categorySheet) {
            sendErrorMessage(chatId, CONFIG.ERROR_SHEET_NOT_FOUND, CONFIG);
            return;
          }
        
          // Get both sheet names and display names
          const subcategoryData = categorySheet.getRange('A2:B' + categorySheet.getLastRow())
            .getValues();
        
          // Process subcategories with display names
          const validSubcategories = subcategoryData
            .map(row => ({
              sheetName: String(row[0]).trim(),
              displayName: row[1] ? String(row[1]).trim() : String(row[0]).trim()
            }))
            .filter(entry => {
              const hasValidFormat = entry.sheetName.includes('🔸');
              const sheetExists = ss.getSheetByName(entry.sheetName);
              return hasValidFormat && sheetExists;
            });
        
          // Create complete entries
          const subcategoryEntries = validSubcategories.map(entry => ({
            sheet: ss.getSheetByName(entry.sheetName),
            displayName: entry.displayName
          }));
        
          sendSubcategoryMenu(chatId, messageId, subcategoryEntries, query, CONFIG, page);
          break;
        

      case 'sheet':
        sendMainMenu(chatId, ss, CONFIG, page);
        break;

      case 'name':
        sendNameMenu(chatId, messageId, ss, query, page, CONFIG);
        break;

      case 'link_sheet':
        sendLinkedSheetsMenu(chatId, messageId, ss, query, CONFIG, page);
        break;

      default:
        sendErrorMessage(chatId, `Invalid pagination type: ${type}`, CONFIG);
        break;
    }
  } catch (error) {
    sendErrorMessage(chatId, `Operation failed: ${error.message}`, CONFIG);
  }
}

// Processes regular text commands
function handleRegularCommands(text, chatId, ss, CONFIG) {
  if (/^\/categories|menu$/i.test(text)) {
    sendMainMenu(chatId, ss, CONFIG);
  } else if (/^\/help$/i.test(text)) {
    sendHelp(chatId, CONFIG);
  } else {
    sendHelp(chatId, CONFIG);
  }

  if (/^\/search$/i.test(text)) {
    sendTelegramMessage({
        method: 'sendMessage',
        chat_id: chatId,
        text: CONFIG.TEXT_SEARCH_PROMPT,
        reply_markup: { force_reply: true }
    }, CONFIG);
    return;
  }

}
