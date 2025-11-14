// ------------------------------------------------------------------ Configuration & Constants
// Your BotToken
var token = "*****";

// Users sheet name
var sheetUsers = "*****";

// Tickets Sheet name
var sheetTickets = "*****";

// Suggestions sheet name
var sheetSuggestions = "*****";

// Definition of ticket table columns - Other functions will be used to detect columns
// The following is for documentation only, actual column detection occurs at runtime


const TICKET_COLUMNS = {
  TIMESTAMP: null,               // Timestamp
  EMAIL: null,                   // Email address
  TRAVELER_NAME: null,           // Traveler(s) Name
  DEPARTURE_LOCATION: null,      // Departure
  ARRIVAL_LOCATION: null,        // Arrival
  TICKET_TYPE: null,             // Ticket Type
  DEPARTURE_DATE: null,          // Departure Date and Time
  RETURN_DATE: null,             // Return Date and Time (Optional for Round Trip)
  TICKET_ID: null,               // Ticket ID
  EMPLOYEE_OPERATIONS: null,     // Employee Name - Operations
  EMPLOYEE_SALES: null,          // Employee Name - Sales
  PURCHASE_FROM: null,           // Purchase From
  PURCHASE_VALUE: null,          // Purchase Value
  SOLD_TO: null,                 // Sold To
  SOLD_VALUE: null,              // Sold Value
  Passport: null,                // Passport
  EDIT: null,                    // Edit
  STATUS: null                   // Status
};

/**
 * Initialization function - Called when the bot starts to detect ticket table columns
 */
function initializeTicketColumns() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
    if (!sheet) {
      Logger.log("Error: Ticket sheet not found");
      return;
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Column detection based on headers
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i]).trim();
      
      if (header.includes("Timestamp")) TICKET_COLUMNS.TIMESTAMP = i;
      else if (header.includes("Email")) TICKET_COLUMNS.EMAIL = i;
      else if (header.includes("Traveler")) TICKET_COLUMNS.TRAVELER_NAME = i;
      else if (header.includes("Departure") && !header.includes("Date")) TICKET_COLUMNS.DEPARTURE_LOCATION = i;
      else if (header.includes("Arrival")) TICKET_COLUMNS.ARRIVAL_LOCATION = i;
      else if (header.includes("Ticket Type")) TICKET_COLUMNS.TICKET_TYPE = i;
      else if (header.includes("Departure Date")) TICKET_COLUMNS.DEPARTURE_DATE = i;
      else if (header.includes("Return Date")) TICKET_COLUMNS.RETURN_DATE = i;
      else if (header.includes("Ticket ID")) TICKET_COLUMNS.TICKET_ID = i;
      else if (header.includes("Employee") && header.includes("Operations")) TICKET_COLUMNS.EMPLOYEE_OPERATIONS = i;
      else if (header.includes("Employee") && header.includes("Sales")) TICKET_COLUMNS.EMPLOYEE_SALES = i;
      else if (header.includes("Purchase From")) TICKET_COLUMNS.PURCHASE_FROM = i;
      else if (header.includes("Purchase Value")) TICKET_COLUMNS.PURCHASE_VALUE = i;
      else if (header.includes("Sold To")) TICKET_COLUMNS.SOLD_TO = i;
      else if (header.includes("Sold Value")) TICKET_COLUMNS.SOLD_VALUE = i;
      else if (header.includes("Passport")) TICKET_COLUMNS.Passport = i;
      else if (header.includes("Edit")) TICKET_COLUMNS.EDIT = i;
      else if (header.includes("Status")) TICKET_COLUMNS.STATUS = i;
    }
    
    // التحقق من اكتشاف الأعمدة الأساسية
    var missingColumns = [];
    if (TICKET_COLUMNS.TIMESTAMP === null) missingColumns.push("Timestamp");
    if (TICKET_COLUMNS.TICKET_ID === null) missingColumns.push("Ticket ID");
    if (TICKET_COLUMNS.STATUS === null) missingColumns.push("Status");
    
    if (missingColumns.length > 0) {
      Logger.log("Warning: Some critical columns not found: " + missingColumns.join(", "));
    } else {
      Logger.log("Successfully initialized ticket columns");
    }
  } catch (error) {
    Logger.log("Error initializing ticket columns: " + error.message);
  }
}

// User Session Manager
var userSessionManager = {
  sessions: {},
  
  // createSession
  createSession: function(chatId, context) {
    this.sessions[chatId] = {
      chatId: chatId,
      context: context || {},
      lastActivity: new Date().getTime()
    };
    return this.sessions[chatId];
  },
  
  // get or create new Session
  getSession: function(chatId) {
    if (!this.sessions[chatId]) {
      return this.createSession(chatId);
    }
    
    // update activity
    this.sessions[chatId].lastActivity = new Date().getTime();
    return this.sessions[chatId];
  },
  
  // updateContext
  updateContext: function(chatId, contextData) {
    const session = this.getSession(chatId);
    session.context = Object.assign({}, session.context, contextData);
    return session;
  },
  
  // removeFromContext
  removeFromContext: function(chatId, keys) {
    const session = this.getSession(chatId);
    if (!session.context) return session;
    
    if (Array.isArray(keys)) {
      keys.forEach(key => delete session.context[key]);
    } else {
      delete session.context[keys];
    }
    
    return session;
  },
  
  // clearSession
  clearSession: function(chatId) {
    delete this.sessions[chatId];
  },
  
  // cleanupSessions
  cleanupSessions: function(maxAgeMs = 30 * 60 * 1000) { // 30 دقيقة افتراضيًا
    const now = new Date().getTime();
    Object.keys(this.sessions).forEach(chatId => {
      if (now - this.sessions[chatId].lastActivity > maxAgeMs) {
        this.clearSession(chatId);
      }
    });
  }
};

var sheetUsers1 = "Users1";

// ------------------------------------------------------------------ Text Handling  System

// Function to process user text and make comparisons case-insensitive
function processText(text) {
  if (!text) return "";

  // trimer
  let processed = text.trim();

  // Command Map that links alternative words to official commands
  const commandMap = {
    'start': ['/start', 'start', 'بدء', 'ابدأ'],
    'help': ['/help', 'help', 'مساعدة', 'مساعده'],
    'add': ['/add', 'add', 'اضافة', 'إضافة', 'إضافه', 'اضافه'],
    'tickets': ['/tickets', 'tickets', 'تذاكر', 'تذكرة', 'تذكره'],
    'search': ['/search', 'search', 'بحث'],
    'analytics': ['/analytics', 'analytics', 'احصائيات', 'إحصائيات', 'تحليلات'],
    'yes': ['yes', 'نعم', 'موافق', 'اي', 'y'],
    'users': ['/users', 'users', 'user management', 'إدارة المستخدمين', 'ادارة المستخدمين', 'user mang'],
    'suggestions': ['/suggestions', 'suggestions', 'اقتراحات', 'إقتراحات']
  };

  // map Checker
  for (const [standard, alternatives] of Object.entries(commandMap)) {
    if (alternatives.some(alt => processed.toLowerCase() === alt.toLowerCase())) {
      return standard; // standardized
    }
  }


  return processed;
}

// ------------------------------------------------------------------ Authorization System

// Check if the user is allowed to use the bot (listed in the users table)

function isAuthorized(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers);
  var data = sheet.getDataRange().getValues();
  // Check if chatId exists in column B (index 1)
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1]) === String(chatId)) {
      return true;
    }
  }
  return false;
}

// Check if the user has "admin" permissions


function isAdmin(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    // Column B (index 1) = chatId, Column D (index 3) = Permission_Level
    if (String(data[i][1]) === String(chatId) && String(data[i][3]).trim() === "مشرف") {
      return true;
    }
  }

  return false;
}


// ------------------------------------------------------------------ Telegram Core System - Event-Based

// سجل الأوامر commandRegistry - يتم تسجيل كل الأوامر ومعالجاتها هنا
const commandRegistry = {
  handlers: {},
  
  // تسجيل معالج أمر جديد
  register: function(command, handler) {
    this.handlers[command] = handler;
    return this; // للسماح بتسلسل التسجيلات
  },
  
  // الحصول على المعالج المناسب للأمر
  getHandler: function(command) {
    return this.handlers[command] || null;
  },
  
  // تنفيذ الأمر
  execute: function(command, chatId, params) {
    const handler = this.getHandler(command);
    if (handler) {
      try {
        handler(chatId, params);
        return true;
      } catch (error) {
        Logger.log(`Error executing command ${command}: ${error.message}`);
        sendMessage(chatId, `❌ حدث خطأ أثناء تنفيذ الأمر: ${error.message}`);
        return false;
      }
    }
    return false;
  }
};

// سجل معالجي الردود الخاصة - للتعامل مع الردود على أسئلة محددة
const responseHandlerRegistry = {
  getHandler: function(chatId) {
    // الحصول على النوع الحالي للرد المتوقع من المستخدم
    const session = userSessionManager.getSession(chatId);
    const { waitingFor } = session.context;
    
    if (!waitingFor) return null;
    
    // إرجاع المعالج المناسب بناءً على نوع الرد المنتظر
    switch (waitingFor) {
      case 'search_term':
        return processSearchTerm;
      case 'add_main_user_id':
        return processAddMainUserStep1;
      case 'add_main_user_name':
        return processAddMainUserStep2;
      case 'add_broadcast_user_id':
        return processAddBroadcastUserStep1;
      case 'add_broadcast_user_name':
        return processAddBroadcastUserStep2;
      case 'edit_main_user_name':
        return processEditMainUserName;
      case 'add_suggestion_value':
        return processAddSuggestionValue;
      case 'edit_suggestion_value':
        return processEditSuggestionValue;
      default:
        return null;
    }
  }
};

// سجل معالجي طلبات الردود (callback queries)
const callbackQueryRegistry = {
  handlers: {},
  
  // تسجيل معالج طلب رد جديد
  register: function(pattern, handler) {
    this.handlers[pattern] = handler;
    return this;
  },
  
  // تنفيذ معالج طلب الرد
  execute: function(chatId, callbackData, message) {
    // التحقق من المعالجات بترتيب التسجيل
    for (const [pattern, handler] of Object.entries(this.handlers)) {
      // إذا كان النمط نصًا دقيقًا
      if (pattern === callbackData) {
        handler(chatId, callbackData, message);
        return true;
      }
      
      // إذا كان النمط يبدأ بـ
      if (pattern.endsWith('*') && callbackData.startsWith(pattern.slice(0, -1))) {
        handler(chatId, callbackData, message);
        return true;
      }
    }
    
    // لم يتم العثور على معالج مناسب
    Logger.log(`No handler found for callback data: ${callbackData}`);
    return false;
  }
};

// الدالة الأساسية التي يتم استدعاؤها عند استقبال أي رسالة من Telegram
function doPost(e) {
  // تهيئة أعمدة التذاكر عند استقبال أي طلب
  initializeTicketColumns();
  
  // التحقق من وجود بيانات
  if (!e || !e.postData || !e.postData.contents) {
    Logger.log("No post data received");
    return HtmlService.createHtmlOutput("No data");
  }

  var data = JSON.parse(e.postData.contents);
  var message = data.message;
  var callbackQuery = data.callback_query;

  try {
    if (message) {
      var chatId = message.chat.id;
      var text = message.text || "";

      Logger.log("Received message: " + text);

      // التحقق من السماح للمستخدم باستخدام البوت
      if (!isAuthorized(chatId)) {
        sendMessage(chatId, "⚠️ غير مسموح لك بالدخول. 😞");
        return HtmlService.createHtmlOutput("Unauthorized");
      }

      // التحقق مما إذا كان المستخدم ينتظر ردًا محددًا (مثل البحث، إضافة مستخدم، إلخ)
      const session = userSessionManager.getSession(chatId);
      const responseHandler = responseHandlerRegistry.getHandler(chatId);
      
      if (responseHandler) {
        // معالجة الرد الخاص
        responseHandler(chatId, text);
        return HtmlService.createHtmlOutput("Response handled");
      }

      // معالجة الأوامر العادية
      var processedText = processText(text);
      const commandExecuted = commandRegistry.execute(processedText, chatId, { message });
      
      // إذا لم يكن أمرًا معروفًا ولم يكن المستخدم في محادثة نشطة، نقوم بتشغيل البحث
      if (!commandExecuted) {
        if (session.context.waitingFor) {
          // المستخدم في محادثة نشطة، لكن لم يتم معالجتها بواسطة responseHandler
          Logger.log(`User in active conversation (${session.context.waitingFor}) but no handler matched`);
        } else {
          // إذا لم يكن أمرًا معروفًا، نفترض أنه بحث
          processSearchTerm(chatId, text);
        }
      }
    }

    // معالجة ردود الأزرار (callbackQuery)
    if (callbackQuery) {
      var chatId = callbackQuery.message.chat.id;
      var callbackData = callbackQuery.data;
      
      Logger.log("Received callback query: " + callbackData);
      
      callbackQueryRegistry.execute(chatId, callbackData, callbackQuery.message);
    }
  } catch (error) {
    Logger.log(`Error in doPost: ${error.message}`);
  }

  return HtmlService.createHtmlOutput("OK");
}

// ------------------------------------------------------------------ Register Commands

commandRegistry
  .register("start", function(chatId) {
    if (isAdmin(chatId)) {
      sendMessage(chatId,
        "🌟 *مرحباً بك في بوت نظام إدارة التذاكر* 🌟\n\n" +
        "👤 *أوامر المستخدمين الأساسية:*\n" +
        "├── /start - عرض شاشة الترحيب وتحديث البوت 🏠  \n" +
        "├── /add - إنشاء تذكرة جديدة 🎟️  \n" +
        "├── /tickets - عرض تذاكرك المفتوحة 📄  \n" +
        "├── /search - البحث في التذاكر باسم العميل أو المكتب أو رقم التذكرة🔍  \n" +
        "└── /help - (فيديو) الدليل المساعد 💬\n\n" +
        "🔐 *أوامر المشرفين المتقدمة:*\n" +
        "├── /analytics - عرض الإحصائيات الشاملة 📈  \n" +
        "└── /users -  إدارة صلاحيات المستخدمين وتحديد قائمة المراد تنبيههم 👥  \n"
      );
    } else {
      sendMessage(chatId,
        "🌟 *مرحباً بك في بوت نظام إدارة التذاكر* 🌟\n\n" +
        "📋 *  الأوامر المتاحة لديك - كمستخدم عادي :*\n" +
        "├── /start -  تحديث البوت و عرض شاشة البدء 🏠  \n" +
        "├── /add - إنشاء تذكرة دعم جديدة 🎫  \n" +
        "├── /tickets - عرض تذاكرك النشطة 📂 \n" +
        "├── /search - البحث في التذاكر المغلقة والحالية 🔎  \n" +
        "└── /help - (فيديو) الدليل المساعد 💬  \n\n"
      );
    }
  })
  .register("help", function(chatId) {
    if (isAdmin(chatId)) {
      sendMessage(chatId,
        "📚 *دليل الأوامر الكامل للمشرفين* 📚\n\n" +
        "👤 *أوامر المستخدمين:*\n" +
        "├── /add - إضافة تذكرة جديدة [شاهد الشرح]\n (https://example.com/add-admin-guide)\n" +
        "├── /tickets - عرض جميع التذاكر المفتوحة [شاهد الشرح]\n (https://example.com/tickets-admin-guide)\n" +
        "└── /search - بحث متقدم (اسم/مكتب/رقم التذكرة) [شاهد الشرح] \n (https://example.com/search-admin-guide)\n\n" +
        "🛠️ *أوامر الإدارة:*\n" +
        "├── /analytics - إحصائيات الأداء والتقارير 📊 [شاهد الشرح] \n (https://example.com/analytics-guide)\n" +
        "└── /users - إضافة/حذف/تعديل صلاحيات المستخدمين 👥 [شاهد الشرح] \n (https://example.com/users-guide)\n" +
        "└── /help -  \n عرض هذه الرسالة المساعدة ❓ [فيديو توضيحي عام](https://example.com/help-how-to)\n\n"+
        "  \n لو تحتاج أي توضيح إضافي: تواصل مباشرة: @mayarIbrahim143 \n  "
      );
    } else {
      sendMessage(chatId,
        "📖 *الدليل المساعد للمستخدمين* 📖\n\n" +
        "🔧 *كيفية استخدام الأوامر:*\n" +
        "├── /add -    \n إنشاء تذكرة دعم جديدة [فيديو توضيحي](https://example.com/add-how-to)\n" +
        "│   (اكتب الأمر ثم اتبع الخطوات البسيطة لإضافة تذكرتك)\n" +
        "├── /tickets -  \n عرض جميع تذاكرك المفتوحة وإدارتها [فيديو توضيحي](https://example.com/tickets-how-to)\n" +
        "├── /search -  \n البحث السريع في التذاكر حتى المغلق منها [فيديو توضيحي](https://example.com/search-how-to)\n" +
        "│   (يمكنك البحث باسم العميل، رقم التذكرة، أو المكتب)\n" +
        "└── /help -  \n عرض هذه الرسالة المساعدة ❓ [فيديو توضيحي عام](https://example.com/help-how-to)\n\n"+
                "\n  لو تحتاج أي توضيح إضافي: تواصل مباشرة: @mayarIbrahim143 \n "

      );
    }
  })  .register("add", function(chatId) {
    startTicketConversation(chatId);
  })
  .register("tickets", function(chatId) {
    showMonthSelection(chatId);
  })
  .register("search", function(chatId) {
    initiateSearch(chatId);
  })
  .register("analytics", function(chatId) {
    if (isAdmin(chatId)) {
      showAnalyticsDashboard(chatId);
    } else {
      sendMessage(chatId, "⚠️ ليس لديك صلاحية الوصول إلى الإحصائيات.");
    }
  })
  .register("users", function(chatId) {
    if (isAdmin(chatId)) {
      showUserManagementMenu(chatId);
    } else {
      sendMessage(chatId, "⚠️ ليس لديك صلاحية إدارة المستخدمين.");
    }
  })
  .register("suggestions", function(chatId) {
    if (isAdmin(chatId)) {
      showSuggestionsMenu(chatId);
    } else {
      sendMessage(chatId, "⚠️ ليس لديك صلاحية إدارة الاقتراحات.");
    }
  });

// ------------------------------------------------------------------ Register Callback Query Handlers

// تسجيل معالجات طلبات الردود
callbackQueryRegistry
  // إدارة المستخدمين
  .register("user_management_main", function(chatId) {
    showUserManagementMenu(chatId);
  })
  .register("back_to_user_management", function(chatId) {
    showUserManagementMenu(chatId);
  })
  .register("user_manage_main", function(chatId) {
    showMainUsersManagement(chatId);
  })
  .register("user_manage_broadcast", function(chatId) {
    showBroadcastUsersManagement(chatId);
  })
  .register("add_main_user", function(chatId) {
    startAddMainUser(chatId);
  })
  .register("add_broadcast_user", function(chatId) {
    startAddBroadcastUser(chatId);
  })
  .register("list_main_users", function(chatId) {
    listMainUsers(chatId);
  })
  .register("list_broadcast_users", function(chatId) {
    listBroadcastUsers(chatId);
  })
  .register("add_user_permission_*", function(chatId, callbackData) {
    var permission = callbackData.split('_')[3];
    var session = userSessionManager.getSession(chatId);
    var { userId, name } = session.context;
    
    if (userId && name) {
      addMainUser(chatId, userId, name, permission);
      // مسح بيانات السياق بعد الإضافة
      userSessionManager.removeFromContext(chatId, ['userId', 'name']);
    } else {
      sendMessage(chatId, "❗️ بيانات المستخدم غير مكتملة. يرجى المحاولة مرة أخرى.");
    }
  })
  .register("edit_main_user_*", function(chatId, callbackData) {
    var userId = callbackData.replace("edit_main_user_", "");
    
    var success = toggleUserPermission(userId);
    if (success) {
      listMainUsers(chatId);
    } else {
      sendMessage(chatId, "❗️ حدث خطأ أثناء تغيير الصلاحية.");
    }
  })
  .register("change_permission_*", function(chatId, callbackData) {
    var parts = callbackData.split('_');
    var userId = parts[2];
    var permission = parts[3];
    changeMainUserPermission(chatId, userId, permission);
  })
  .register("delete_main_user_*", function(chatId, callbackData) {
    var userId = callbackData.split('_')[3];
    deleteMainUser(chatId, userId);
    setTimeout(function() {
      listMainUsers(chatId);
    }, 1000);
  })
  .register("delete_broadcast_user_*", function(chatId, callbackData) {
    var userId = callbackData.split('_')[3];
    deleteBroadcastUser(chatId, userId);
    setTimeout(function() {
      listBroadcastUsers(chatId);
    }, 1000);
  })
  // مزيد من الأزرار
  .register("show_analytics", function(chatId) {
    showAnalyticsDashboard(chatId);
  })
  .register("analytics_daily", function(chatId) {
    showDailyTrends(chatId);
  })
  .register("analytics_employees", function(chatId) {
    Logger.log("DEBUG: analytics_employees callback triggered for chat ID: " + chatId);
    try {
      showEmployeeAnalysis(chatId);
    } catch (error) {
      Logger.log("ERROR in analytics_employees: " + error.message);
      Logger.log(error.stack);
      sendMessage(chatId, "❌ حدث خطأ أثناء تحليل أداء الموظفين: " + error.message);
    }
  })
  .register("analytics_export", function(chatId) {
    Logger.log("DEBUG: analytics_export callback triggered for chat ID: " + chatId);
    try {
      exportAnalyticsToExcel(chatId);
    } catch (error) {
      Logger.log("ERROR in analytics_export: " + error.message);
      Logger.log(error.stack);
      sendMessage(chatId, "❌ حدث خطأ أثناء تصدير البيانات: " + error.message);
    }
  })
  .register("back_to_main", function(chatId) {
    commandRegistry.execute("start", chatId);
  })
  .register("search_field_*", function(chatId, callbackData) {
    var parts = callbackData.split("_");
    var field = parts[2];
    var term = decodeURIComponent(parts.slice(3).join("_"));
    processSearchTermByField(chatId, term, field);
  })
  .register("view_ticket_*", function(chatId, callbackData) {
    var parts = callbackData.split('_');
    var ticketId = parts[2];
    var searchTerm = decodeURIComponent(parts.slice(3).join('_'));
    displayTicketDetails(chatId, ticketId, searchTerm);
  })
  .register("search_results_*", function(chatId, callbackData) {
    var searchTerm = decodeURIComponent(callbackData.split('_')[2]);
    returnToSearchResults(chatId, searchTerm);
  })
  .register("close_search_ticket_*", function(chatId, callbackData) {
    var parts = callbackData.split('_');
    var ticketId = parts[3];
    var searchTerm = decodeURIComponent(parts.slice(4).join('_'));
    closeTicketFromSearch(chatId, ticketId, searchTerm);
  })
  .register("month_*", function(chatId, callbackData) {
    var monthKey = callbackData.split("_")[1];
    showTicketsForMonth(chatId, monthKey);
  })
  .register("ticket_*", function(chatId, callbackData, message) {
    var parts = callbackData.split('_');
    
    // تأكد أن البيانات فيها ticketId وmonthKey
    if (parts.length >= 3) {
      var ticketId = parts[1];
      var monthKey = parts[2];
      showTicketDetails(chatId, ticketId, monthKey);
    } else {
      sendMessage(chatId, "❗️ بيانات التذكرة غير صحيحة.");
    }
  })
  .register("close_ticket_*", function(chatId, callbackData) {
    // تقسيم البيانات للحصول على معرف التذكرة والشهر
    var parts = callbackData.split('_');
    
    if (parts.length >= 4) {
      var monthKey = parts[2];
      var ticketId = parts[3];
      
      // التحقق من أن المستخدم مشرف
      if (!isAdmin(chatId)) {
        sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم إغلاق التذاكر.");
        return;
      }
      
      // إغلاق التذكرة
      closeTicket(chatId, ticketId);
      
      // بعد الإغلاق، العودة لقائمة تذاكر الشهر
      setTimeout(function() {
        showTicketsForMonth(chatId, monthKey);
      }, 1000);
    } else {
      sendMessage(chatId, "❗️ بيانات إغلاق التذكرة غير صحيحة.");
    }
  })
  .register("back_to_month_*", function(chatId, callbackData) {
    var monthKey = callbackData.replace("back_to_month_", "");
    showTicketsForMonth(chatId, monthKey);
  })
  .register("show_suggestions", function(chatId) {
    showSuggestionsMenu(chatId);
  })
  .register("suggestions_header_*", function(chatId, callbackData) {
    var headerIndex = parseInt(callbackData.split('_')[2]);
    showSuggestionsColumn(chatId, headerIndex);
  })
  .register("add_suggestion_for_*", function(chatId, callbackData) {
    var headerIndex = parseInt(callbackData.split('_')[3]);
    startAddSuggestion(chatId, headerIndex);
  })
  .register("suggestion_value_*", function(chatId, callbackData) {
    var parts = callbackData.split('_');
    var headerIndex = parseInt(parts[2]);
    var valueIndex = parseInt(parts[3]);
    showSuggestionValueOptions(chatId, headerIndex, valueIndex);
  })
  .register("edit_suggestion_*_*", function(chatId, callbackData) {
    var parts = callbackData.split('_');
    var headerIndex = parseInt(parts[2]);
    var valueIndex = parseInt(parts[3]);
    startEditSuggestion(chatId, headerIndex, valueIndex);
  })
  .register("delete_suggestion_*_*", function(chatId, callbackData) {
    var parts = callbackData.split('_');
    var headerIndex = parseInt(parts[2]);
    var valueIndex = parseInt(parts[3]);
    deleteSuggestion(chatId, headerIndex, valueIndex);
  })
  .register("back_to_suggestions", function(chatId) {
    showSuggestionsMenu(chatId);
  })
  .register("back_to_suggestion_column_*", function(chatId, callbackData) {
    var headerIndex = parseInt(callbackData.split('_')[4]);
    showSuggestionsColumn(chatId, headerIndex);
  })
  // إضافة معالجات فترات التحليل
  .register("analytics_period_this_day", function(chatId) {
    analyzeTicketsForPeriod(chatId, "this_day");
  })
  .register("analytics_period_last_day", function(chatId) {
    analyzeTicketsForPeriod(chatId, "last_day");
  })
  .register("analytics_period_this_week", function(chatId) {
    analyzeTicketsForPeriod(chatId, "this_week");
  })
  .register("analytics_period_last_week", function(chatId) {
    analyzeTicketsForPeriod(chatId, "last_week");
  })
  .register("analytics_period_this_month", function(chatId) {
    analyzeTicketsForPeriod(chatId, "this_month");
  })
  .register("analytics_period_last_month", function(chatId) {
    analyzeTicketsForPeriod(chatId, "last_month");
  })
  .register("analytics_period_this_quarter", function(chatId) {
    analyzeTicketsForPeriod(chatId, "this_quarter");
  })
  .register("analytics_period_last_quarter", function(chatId) {
    analyzeTicketsForPeriod(chatId, "last_quarter");
  })
  .register("analytics_period_this_year", function(chatId) {
    analyzeTicketsForPeriod(chatId, "this_year");
  })
  .register("analytics_period_last_year", function(chatId) {
    analyzeTicketsForPeriod(chatId, "last_year");
  })
  .register("analytics_period_all", function(chatId) {
    analyzeTicketsForPeriod(chatId, "all");
  })
  .register("analytics_employees_*", function(chatId, callbackData) {
    var period = callbackData.replace("analytics_employees_", "");
    Logger.log("DEBUG: analytics_employees_* callback triggered with period: " + period);
    try {
      showEmployeeAnalysis(chatId, period);
    } catch (error) {
      Logger.log("ERROR in analytics_employees_*: " + error.message);
      Logger.log(error.stack);
      sendMessage(chatId, "❌ حدث خطأ أثناء تحليل أداء الموظفين: " + error.message);
    }
  })
  .register("analytics_export_*", function(chatId, callbackData) {
    var period = callbackData.replace("analytics_export_", "");
    Logger.log("DEBUG: analytics_export_* callback triggered with period: " + period);
    try {
      exportAnalyticsToExcel(chatId, period);
    } catch (error) {
      Logger.log("ERROR in analytics_export_*: " + error.message);
      Logger.log(error.stack);
      sendMessage(chatId, "❌ حدث خطأ أثناء تصدير البيانات: " + error.message);
    }
  })
  .register("toggle_ticket_status_*", function(chatId, callbackData) {
    var parts = callbackData.split('_');
    var ticketId = parts[3];
    var searchTerm = decodeURIComponent(parts.slice(4).join('_'));
    toggleTicketStatus(chatId, ticketId, searchTerm);
  })
  .register("edit_ticket_*", function(chatId, callbackData) {
    var parts = callbackData.split('_');
    var ticketId = parts[2];
    var searchTerm = parts.length > 3 ? decodeURIComponent(parts.slice(3).join('_')) : null;
    
    if (isAdmin(chatId)) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
      var data = sheet.getDataRange().getValues();
      
      // Initialize ticket columns if not already initialized
      if (TICKET_COLUMNS.EDIT === null) {
        initializeTicketColumns();
      }
      
      var ticket = null;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][TICKET_COLUMNS.TICKET_ID]) == String(ticketId)) {
          ticket = {
            editLink: data[i][TICKET_COLUMNS.EDIT]
          };
          break;
        }
      }

      if (ticket && ticket.editLink) {
        sendMessage(chatId, `🔗 يمكنك تعديل التذكرة عبر الرابط التالي: ${ticket.editLink}`);
        // إذا كان هناك مصطلح بحث، عُد إلى تفاصيل التذكرة بعد إرسال الرابط
        if (searchTerm) {
          setTimeout(function() {
            displayTicketDetails(chatId, ticketId, searchTerm);
          }, 1000);
        }
      } else {
        sendMessage(chatId, "❌ الرابط غير متاح للتعديل.");
      }
    } else {
      sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم تعديل التذكرة.");
    }
  });

// ------------------------------------------------------------------ Message Sending System

// إرسال رسالة إلى المستخدم في Telegram
function sendMessage(chatId, text, replyMarkup = null, replyToMessageId = null) {
  var url = "https://api.telegram.org/bot" + token + "/sendMessage";
  
  var payload = {
    chat_id: chatId,
    text: text,
    parse_mode: "HTML"
  };

  if (replyMarkup) {
    payload.reply_markup = JSON.stringify(replyMarkup);
  }
  
  if (replyToMessageId) {
    payload.reply_to_message_id = replyToMessageId;
  }

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseData = JSON.parse(response.getContentText());
    return responseData.ok ? responseData.result.message_id : null;
  } catch (e) {
    Logger.log("Error sending message: " + e.message);
    return null;
  }
}

// تعديل الدوال المتعلقة بالبحث والإجراءات الأخرى لاستخدام مدير الجلسات الجديد

// تهيئة البحث
function initiateSearch(chatId) {
    // Clear any existing state and set to search mode
    userSessionManager.updateContext(chatId, { waitingFor: 'search_term' });
    
    // Create a force reply markup to ensure the bot knows the next message is a reply
    var forceReplyMarkup = {
      force_reply: true,
      selective: true
    };
    
    // Send message with force reply
    sendMessage(
      chatId, 
      "🔍 الرجاء إرسال اسم المسافر للبحث:", 
      forceReplyMarkup
    );
    
    Logger.log("Search initiated for chatId: " + chatId + " with force reply");
  }

// Handle search term from user and perform search
function processSearchTerm(chatId, searchTerm) {
  Logger.log("Processing search term: '" + searchTerm + "' for chatId: " + chatId);

  if (!searchTerm || searchTerm.trim() === "") {
    sendMessage(chatId, "⚠️ الرجاء إدخال اسم صحيح للبحث.");
    userSessionManager.removeFromContext(chatId, 'waitingFor');
    return;
  }

  // جلب البيانات
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0]; // Get headers
  var ticketRows = rows.slice(1); // Skip header row
  var norm = searchTerm.toLowerCase().trim();

  // Initialize ticket columns if not already initialized
  if (TICKET_COLUMNS.TRAVELER_NAME === null) {
    initializeTicketColumns();
  }

  // بحث في أسماء المسافرين
  var results = ticketRows.filter(row => {
    var travelerName = row[TICKET_COLUMNS.TRAVELER_NAME] || "";
    return travelerName.toString().toLowerCase().includes(norm);
  }).map((row, idx) => ({
    rowIndex: idx + 2,
    ticketId: row[TICKET_COLUMNS.TICKET_ID] || "N/A",
    purchaseFrom: row[TICKET_COLUMNS.PURCHASE_FROM] || "N/A",
    soldTo: row[TICKET_COLUMNS.SOLD_TO] || "N/A",
    travelerName: row[TICKET_COLUMNS.TRAVELER_NAME] || "N/A",
    departureDate: row[TICKET_COLUMNS.DEPARTURE_DATE] || "N/A",
    status: row[TICKET_COLUMNS.STATUS] || "N/A"
  }));

  if (results.length === 0) {
    // لا توجد نتائج: عرض خيارات حقول البحث
    var keyboard = [
      [{ text: "🏢 بحث في Purchase From", callback_data: `search_field_purchase_${encodeURIComponent(searchTerm)}` }],
      [{ text: "🏢 بحث في Sold To", callback_data: `search_field_sold_${encodeURIComponent(searchTerm)}` }],
      [{ text: "🎫 بحث برقم التذكرة", callback_data: `search_field_ticket_${encodeURIComponent(searchTerm)}` }]
    ];

    sendMessage(
      chatId,
      `🔍 لم يتم العثور على "${searchTerm}" في أسماء المسافرين. هل ترغب في البحث في مجال آخر؟`,
      { inline_keyboard: keyboard }
    );

    userSessionManager.updateContext(chatId, { waitingFor: 'choose_search_field', searchTerm: searchTerm });
    return;
  }

  // إذا وُجدت نتائج: عرضها مباشرة
  displaySearchResults(chatId, results, searchTerm);
}

function processSearchTermByField(chatId, searchTerm, field) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0]; // Get headers
  var ticketRows = rows.slice(1); // Skip header row
  var norm = searchTerm.toLowerCase().trim();

  // Initialize ticket columns if not already initialized
  if (TICKET_COLUMNS.TRAVELER_NAME === null) {
    initializeTicketColumns();
  }

  var results = ticketRows.filter(row => {
    var fieldValue = "";
    if (field === "traveler") fieldValue = row[TICKET_COLUMNS.TRAVELER_NAME] || "";
    if (field === "purchase") fieldValue = row[TICKET_COLUMNS.PURCHASE_FROM] || "";
    if (field === "sold") fieldValue = row[TICKET_COLUMNS.SOLD_TO] || "";
    if (field === "ticket") fieldValue = row[TICKET_COLUMNS.TICKET_ID] ? row[TICKET_COLUMNS.TICKET_ID].toString() : "";
    
    return fieldValue.toString().toLowerCase().includes(norm);
  }).map((row, idx) => ({
    rowIndex: idx + 2,
    ticketId: row[TICKET_COLUMNS.TICKET_ID] || "N/A",
    purchaseFrom: row[TICKET_COLUMNS.PURCHASE_FROM] || "N/A",
    soldTo: row[TICKET_COLUMNS.SOLD_TO] || "N/A",
    travelerName: row[TICKET_COLUMNS.TRAVELER_NAME] || "N/A",
    departureDate: row[TICKET_COLUMNS.DEPARTURE_DATE] || "N/A",
    status: row[TICKET_COLUMNS.STATUS] || "N/A"
  }));

  var fieldLabel = field === "traveler" ? "أسماء المسافرين"
                 : field === "purchase" ? "Purchase From"
                 : field === "sold" ? "Sold To"
                 : "أرقام التذاكر";
  
  if (results.length === 0) {
    var keyboard = [
      [{ text: "🏢 بحث في Purchase From", callback_data: `search_field_purchase_${encodeURIComponent(searchTerm)}` }],
      [{ text: "🏢 بحث في Sold To", callback_data: `search_field_sold_${encodeURIComponent(searchTerm)}` }],
      [{ text: "🎫 بحث برقم التذكرة", callback_data: `search_field_ticket_${encodeURIComponent(searchTerm)}` }],
      [{ text: "👥 بحث في اسم المسافر", callback_data: `search_field_traveler_${encodeURIComponent(searchTerm)}` }]
    ];

    sendMessage(
      chatId,
      `⚠️ لا توجد نتائج في ${fieldLabel} لـ "${searchTerm}". هل ترغب في البحث في مجال آخر؟`,
      { inline_keyboard: keyboard }
    );
    
    userSessionManager.updateContext(chatId, { waitingFor: 'choose_search_field', searchTerm: searchTerm });
    return;
  }
  
  displaySearchResults(chatId, results, searchTerm);
}

// ====================================================================
// 📌 وظيفة: عرض نتائج البحث كمجموعة أزرار تفاعلية للمستخدم
function displaySearchResults(chatId, results, searchTerm) {
  // إنشاء لوحة مفاتيح تفاعلية تحتوي على نتائج البحث
  var inlineKeyboard = results.map(r => [{
    text: `${r.status.includes("مفتوحة") || r.status.includes("Open") ? "🟢" : "🔴"} #${r.ticketId} – ${r.travelerName}`,
    callback_data: `view_ticket_${r.ticketId}_${encodeURIComponent(searchTerm)}`
  }]);

  // إرسال الرسالة للمستخدم تتضمن عدد النتائج مع لوحة الأزرار
  sendMessage(
    chatId,
    `🔍 نتائج البحث عن "${searchTerm}" (${results.length}):`,
    { inline_keyboard: inlineKeyboard }
  );
  
  // حذف حالة المستخدم المؤقتة بعد عرض النتائج
  userSessionManager.removeFromContext(chatId, 'waitingFor');
}

// ====================================================================
// 📌 وظيفة: عرض تفاصيل التذكرة عند اختيارها من نتائج البحث
function displayTicketDetails(chatId, ticketId, searchTerm) {
  Logger.log("Displaying details for ticket #" + ticketId);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Initialize ticket columns if not already initialized
  if (TICKET_COLUMNS.TRAVELER_NAME === null) {
    initializeTicketColumns();
  }
  
  var ticket = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][TICKET_COLUMNS.TICKET_ID]) == String(ticketId)) {
      // Get raw display values for dates directly from cells
      var departureDate = sheet.getRange(i + 1, TICKET_COLUMNS.DEPARTURE_DATE + 1).getDisplayValue();
      var returnDate = sheet.getRange(i + 1, TICKET_COLUMNS.RETURN_DATE + 1).getDisplayValue();
      
      ticket = {
        rowIndex: i + 1,
        ticketId: data[i][TICKET_COLUMNS.TICKET_ID],
        purchaseFrom: data[i][TICKET_COLUMNS.PURCHASE_FROM],
        purchaseValue: data[i][TICKET_COLUMNS.PURCHASE_VALUE],
        soldTo: data[i][TICKET_COLUMNS.SOLD_TO],
        soldValue: data[i][TICKET_COLUMNS.SOLD_VALUE],
        travelerName: data[i][TICKET_COLUMNS.TRAVELER_NAME],
        departureLocation: data[i][TICKET_COLUMNS.DEPARTURE_LOCATION],
        arrivalLocation: data[i][TICKET_COLUMNS.ARRIVAL_LOCATION],
        departureDate: departureDate,
        returnDate: returnDate,
        status: data[i][TICKET_COLUMNS.STATUS],
        salesEmployee: data[i][TICKET_COLUMNS.EMPLOYEE_SALES],
        operationsEmployee: data[i][TICKET_COLUMNS.EMPLOYEE_OPERATIONS],
        email: data[i][TICKET_COLUMNS.EMAIL],
        Passport: data[i][TICKET_COLUMNS.Passport],
        editLink: data[i][TICKET_COLUMNS.EDIT]
      };
      break;
    }
  }

  if (!ticket) {
    sendMessage(chatId, "❌ التذكرة غير موجودة أو تم حذفها.");
    return;
  }

  // Direct access to departure date
  var departureDate = ticket.departureDate;
  var formattedDepartureDate = departureDate;
  
  // Direct access to return date
  var returnDate = ticket.returnDate;
  var formattedReturnDate = "";
  
  if (returnDate) {
    formattedReturnDate = returnDate;
  }

  var statusWithEmoji = ticket.status.includes("مفتوحة") || ticket.status.includes("Open") ? "🟢 مفتوحة" : "🔴 مغلقة";

  var ticketDetails = `📋 <b>تفاصيل التذكرة #${ticket.ticketId}</b>\n\n` +
                      `<b>الحالة:</b> ${statusWithEmoji}\n` +
                      `<b>اسم المسافر(ين):</b> ${ticket.travelerName}\n` +
                      `<b>المُدخل:</b> ${ticket.email || "غير متوفر"}\n` +
                      `<b>تاريخ المغادرة:</b> ${formattedDepartureDate}\n` +
                      (returnDate ? `<b>تاريخ العودة:</b> ${formattedReturnDate}\n` : "") +
                      `<b>من:</b> ${ticket.departureLocation}\n` +
                      `<b>إلى:</b> ${ticket.arrivalLocation}\n` +
                      `<b>مصدر الشراء:</b> ${ticket.purchaseFrom}\n` +
                      `<b>سعر الشراء:</b> ${ticket.purchaseValue || "غير متوفر"}\n` +
                      `<b>وجهة البيع:</b> ${ticket.soldTo}\n` +
                      `<b>سعر البيع:</b> ${ticket.soldValue || "غير متوفر"}\n` +
                      `<b>موظف المبيعات:</b> ${ticket.salesEmployee}\n` +
                      `<b>موظف العمليات:</b> ${ticket.operationsEmployee}\n` +
                      `<b>جواز السفر:</b> ${ticket.Passport || "لا توجد"}`;

  // تقسيم الأزرار إلى صفوف
  var row1 = [];
  var row2 = [];

  row1.push({
          text: "🔙 العودة للنتائج", 
          callback_data: `search_results_${encodeURIComponent(searchTerm)}` 
  });

  row1.push({
    text: "🔄 تبديل الحالة",
    callback_data: `toggle_ticket_status_${ticket.ticketId}_${encodeURIComponent(searchTerm)}`
  });

  if (ticket.status.includes("مفتوحة") || ticket.status.includes("Open")) {
    row2.push({
      text: "❌ إغلاق التذكرة",
      callback_data: `close_search_ticket_${ticket.ticketId}_${encodeURIComponent(searchTerm)}`
    });
  }

  if (ticket.editLink) {
    row2.push({
      text: "✏️ تعديل التذكرة",
      callback_data: `edit_ticket_${ticket.ticketId}_${encodeURIComponent(searchTerm)}`
    });
  }

  sendMessage(chatId, ticketDetails, {
    inline_keyboard: [row1, row2]
  });
}

function handleCallbackQuery(chatId, callbackData) {
  // 📝 معالجة الضغط على زر تعديل التذكرة من البحث أو نظام إدارة التذاكر
  if (callbackData.startsWith("edit_ticket_")) {
    var parts = callbackData.split('_');
    var ticketId = parts[2];
    var contextKey = parts.length > 3 ? parts[3] : null; // This could be either monthKey or searchTerm
    var searchTerm = null;
    var monthKey = null;
    
    // Determine if this is from search system or monthly view
    if (contextKey && contextKey.includes('-')) {
      // This is likely a monthKey in format YYYY-MM
      monthKey = contextKey;
    } else if (contextKey) {
      // This is likely a searchTerm
      searchTerm = decodeURIComponent(parts.slice(3).join('_'));
    }
    
    if (isAdmin(chatId)) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
      var data = sheet.getDataRange().getValues();
      
      // Initialize ticket columns if not already initialized
      if (TICKET_COLUMNS.EDIT === null) {
        initializeTicketColumns();
      }
      
      var ticket = null;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][TICKET_COLUMNS.TICKET_ID]) == String(ticketId)) {
          ticket = {
            editLink: data[i][TICKET_COLUMNS.EDIT]
          };
          break;
        }
      }

      if (ticket && ticket.editLink) {
        sendMessage(chatId, `🔗 يمكنك تعديل التذكرة عبر الرابط التالي: ${ticket.editLink}`);
        
        // Navigate back to appropriate context
        setTimeout(function() {
          if (searchTerm) {
            displayTicketDetails(chatId, ticketId, searchTerm);
          } else if (monthKey) {
            showTicketDetails(chatId, ticketId, monthKey);
          }
        }, 1000);
      } else {
        sendMessage(chatId, "❌ الرابط غير متاح للتعديل.");
      }
    } else {
      sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم تعديل التذكرة.");
    }
  }
  
  // معالجة الضغط على أزرار البحث في حقول مختلفة
  else if (callbackData.startsWith("search_field_")) {
    var parts = callbackData.split('_');
    var field = parts[2];
    var searchTerm = decodeURIComponent(parts.slice(3).join('_'));

    // تحويل الحقول إلى المعرفات الجديدة
    if (field === "purchase") {
      processSearchTermByField(chatId, searchTerm, "purchase");
    } 
    else if (field === "sold") {
      processSearchTermByField(chatId, searchTerm, "sold");
    }
    else if (field === "ticket") {
      processSearchTermByField(chatId, searchTerm, "ticket");
    }
    else if (field === "traveler") {
      processSearchTermByField(chatId, searchTerm, "traveler");
    }
  }

  // ✅ معالجة الضغط على زر إغلاق التذكرة
  else if (callbackData.startsWith("close_search_ticket_")) {
    var parts = callbackData.split('_');
    var ticketId = parts[3];
    var searchTerm = decodeURIComponent(parts.slice(4).join('_'));

    if (!isAdmin(chatId)) {
      sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم إغلاق التذاكر.");
      return;
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
    var data = sheet.getDataRange().getValues();
    
    // Initialize ticket columns if not already initialized
    if (TICKET_COLUMNS.STATUS === null) {
      initializeTicketColumns();
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][TICKET_COLUMNS.TICKET_ID]) == String(ticketId)) {
        sheet.getRange(i + 1, TICKET_COLUMNS.STATUS + 1).setValue("مغلقة"); // تحويل العمود من مؤشر مصفوفة إلى مؤشر جدول
        sendMessage(chatId, `✅ تم إغلاق التذكرة رقم ${ticketId}.`);
        displayTicketDetails(chatId, ticketId, searchTerm); // عرض التفاصيل بعد التحديث
        return;
      }
    }

    sendMessage(chatId, "❌ لم يتم العثور على التذكرة.");
  }

  // ✅ معالجة الضغط على زر العودة للنتائج
  else if (callbackData.startsWith("search_results_")) {
    var searchTerm = decodeURIComponent(callbackData.split('_')[2]);
    returnToSearchResults(chatId, searchTerm);
  }

  // ✅ معالجة الضغط على زر "تبديل الحالة" من البحث أو نظام إدارة التذاكر
  else if (callbackData.startsWith("toggle_ticket_status_")) {
    var parts = callbackData.split('_');
    var ticketId = parts[3];
    var contextKey = parts[4]; // This could be either monthKey or searchTerm
    var searchTerm = null;
    var monthKey = null;
    
    // Determine if this is from search system or monthly view
    if (contextKey && contextKey.includes('-')) {
      // This is likely a monthKey in format YYYY-MM
      monthKey = contextKey;
    } else if (contextKey) {
      // This is likely a searchTerm
      searchTerm = decodeURIComponent(parts.slice(4).join('_'));
    }

    if (!isAdmin(chatId)) {
      sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم تبديل حالة التذكرة.");
      return;
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
    var data = sheet.getDataRange().getValues();
    
    // Initialize ticket columns if not already initialized
    if (TICKET_COLUMNS.STATUS === null) {
      initializeTicketColumns();
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][TICKET_COLUMNS.TICKET_ID]) == String(ticketId)) {
        var currentStatus = data[i][TICKET_COLUMNS.STATUS]; // الحصول على الحالة الحالية
        var newStatus;
        
        // تبديل الحالة حسب القيمة الحالية
        if (currentStatus.includes("مفتوحة") || currentStatus.includes("Open")) {
          newStatus = "مغلقة";
        } else {
          newStatus = "مفتوحة";
        }

        // تحديث حالة التذكرة في الجدول
        sheet.getRange(i + 1, TICKET_COLUMNS.STATUS + 1).setValue(newStatus); // تحويل العمود من مؤشر مصفوفة إلى مؤشر جدول
        sendMessage(chatId, `✅ تم تبديل حالة التذكرة رقم ${ticketId} إلى: ${newStatus}`);
        
        // Navigate to appropriate details view
        if (searchTerm) {
          displayTicketDetails(chatId, ticketId, searchTerm);
        } else if (monthKey) {
          showTicketDetails(chatId, ticketId, monthKey);
        }
        return;
      }
    }

    sendMessage(chatId, "❌ لم يتم العثور على التذكرة.");
  }
  
  // معالجة النقر على زر عرض تفاصيل التذكرة
  else if (callbackData.startsWith("view_ticket_")) {
    var parts = callbackData.split('_');
    var ticketId = parts[2];
    var searchTerm = decodeURIComponent(parts.slice(3).join('_'));
    
    displayTicketDetails(chatId, ticketId, searchTerm);
  }
}

// ====================================================================
// 📌 وظيفة: إعادة عرض نتائج البحث بعد العودة من التفاصيل
function returnToSearchResults(chatId, searchTerm) {
  // إعادة تنفيذ البحث لعرض النتائج المحدثة
  processSearchTerm(chatId, searchTerm);
}

// ====================================================================
// 📌 وظيفة: إغلاق التذكرة من نتائج البحث (للمشرف فقط)
function closeTicketFromSearch(chatId, ticketId, searchTerm) {
  // التأكد من أن المستخدم مشرف
  if (!isAdmin(chatId)) {
    sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم إغلاق التذاكر.");
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  // البحث عن الصف الخاص بالتذكرة
  for (var i = 1; i < data.length; i++) {
    if (data[i][TICKET_COLUMNS.TICKET_ID] == ticketId) {
      rowIndex = i + 1;
      break;
    }
  }
  
  // إذا لم يتم العثور على التذكرة
  if (rowIndex === -1) {
    sendMessage(chatId, "❌ التذكرة غير موجودة أو تم حذفها بالفعل.");
    return;
  }
  
  // تحديث الحالة إلى "مغلقة"
  sheet.getRange(rowIndex, TICKET_COLUMNS.STATUS + 1).setValue("مغلقة");
  
  // تأكيد الإغلاق للمستخدم
  sendMessage(chatId, `✅ تم إغلاق التذكرة #${ticketId} بنجاح!`);
  
  // انتظار قصير ثم إعادة عرض نتائج البحث
      Utilities.sleep(1000);
      returnToSearchResults(chatId, searchTerm);
}



// دالة تبديل حالة التذكرة
function toggleTicketStatus(chatId, ticketId, searchTerm) {
  if (!isAdmin(chatId)) {
    sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم تبديل حالة التذكرة.");
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][TICKET_COLUMNS.TICKET_ID] == ticketId) {
      var currentStatus = data[i][TICKET_COLUMNS.STATUS]; // الحصول على الحالة الحالية
      var newStatus = currentStatus === "مفتوحة" ? "مغلقة" : "مفتوحة"; // تبديل الحالة

      // تحديث حالة التذكرة في الجدول
      sheet.getRange(i + 1, TICKET_COLUMNS.STATUS + 1).setValue(newStatus); // العمود F = الحالة
      sendMessage(chatId, `✅ تم تبديل حالة التذكرة رقم ${ticketId} إلى: ${newStatus}`);
      displayTicketDetails(chatId, ticketId, searchTerm); // عرض التفاصيل بعد التحديث
      return;
    }
  }
  
  sendMessage(chatId, "❌ لم يتم العثور على التذكرة.");
}

// ==================================================================== onFormSubmit System
// 📌 اضافة رابط لتعديل التذكرة (Form Submit)
function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tickets");
  var row = e.range.getRow();
  var isEdit = false;

  // Check if this is an edit or a new submission
  // If the edit URL already exists, it's likely an edit
  var ticketId = sheet.getRange(row, 9).getValue() || "غير معروف";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var messageIDsSheet = ss.getSheetByName("MessageIDs");
  
  if (messageIDsSheet) {
    var data = messageIDsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === ticketId) {
        isEdit = true;
        break;
      }
    }
  }
  
  // Fallback check: if the edit URL already exists
  if (!isEdit && sheet.getRange(row, 17).getValue() !== "") {
    isEdit = true;
  }

  // Set status to "مفتوحة" in column R (if it's a new ticket)
  // For edits, we'll keep the existing status
  if (!isEdit) {
    sheet.getRange(row, 18).setValue("مفتوحة");
  }

  // Get the form LINKED TO THIS SHEET
  var form = FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl());
  
  // Get matching response using timestamp
  var timestamp = sheet.getRange(row, 1).getValue(); //  timestamp is in column A
  var response = form.getResponses().find(r => r.getTimestamp().valueOf() === timestamp.valueOf());
  
  if (response) {
    var editUrl = response.getEditResponseUrl();
    sheet.getRange(row, 17).setValue(editUrl); // Add Edit URL to column Q
    Logger.log(" تم تحديث الحالة وإضافة رابط التعديل");
    
    // Current values for storing or comparing
    var currentValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Notify group about the new ticket or edit
    try {
      var mainGroupChatId = "-4609721442"; // Group chat ID
      
      if (isEdit) {
        // For edits, identify and send only the changed fields
        var changedFields = getChangedFields(sheet, row, currentValues);
        
        if (changedFields.length > 0) {
          var editMsg = "✏️ *تم تعديل تذكرة رقم* " + ticketId + "\n\n" + changedFields.join("\n");
          
          // Get the original message ID from MessageIDs sheet
          var originalMsgId = null;
          if (messageIDsSheet) {
            var data = messageIDsSheet.getDataRange().getValues();
            for (var i = 1; i < data.length; i++) {
              if (data[i][0] === ticketId) {
                originalMsgId = data[i][1];
                break;
              }
            }
          }
          
          // If no message ID in sheet, try properties as fallback
          if (!originalMsgId) {
            var props = PropertiesService.getScriptProperties();
            originalMsgId = props.getProperty("ticket_msg_" + ticketId);
          }
          
          if (originalMsgId) {
            // Send as reply to original message
            sendMessage(mainGroupChatId, editMsg, null, originalMsgId);
            // Update stored values
            storeMessageInfo(ticketId, originalMsgId, currentValues);
          } else {
            // If no original message ID, send as regular message
            var messageId = sendMessage(mainGroupChatId, editMsg);
            // Store new message ID and values
            storeMessageInfo(ticketId, messageId, currentValues);
          }
        }
      } else {
        // For new tickets, send the full information
        var ticketInfo = getTicketInfoFromRow(sheet, row, "new");
        
        // Send message and store message ID for future replies
        var messageId = sendMessage(mainGroupChatId, ticketInfo);
        
        // Store message ID and current values
        if (messageId) {
          storeMessageInfo(ticketId, messageId, currentValues);
          
          // Also store in properties as backward compatibility
          var props = PropertiesService.getScriptProperties();
          props.setProperty("ticket_msg_" + ticketId, messageId.toString());
        }
      }
    } catch (error) {
      // If notification fails, notify admin
      var adminChatId = "277264385"; // admin chat ID
      var errorMsg = isEdit ? 
                    "⚠️ فشل في إرسال إشعار تعديل التذكرة:\n" : 
                    "⚠️ فشل في إرسال إشعار التذكرة الجديدة:\n";
      sendMessage(adminChatId, errorMsg + error.toString());
      Logger.log("Error sending notification: " + error.toString());
    }
  } else {
    Logger.log("⚠️ لم يتم العثور على الرد المطابق للتايمسامب");
  }
}

// Helper function to extract ticket information from a row
function getTicketInfoFromRow(sheet, row, type) {
  // Use direct column indices based on the sheet structure
  var email = sheet.getRange(row, 2).getValue() || "غير معروف";           // Email address
  var travelerName = sheet.getRange(row, 3).getValue() || "غير معروف";    // Traveler(s) Name
  var departure = sheet.getRange(row, 4).getValue() || "غير معروف";       // Departure
  var arrival = sheet.getRange(row, 5).getValue() || "غير معروف";         // Arrival
  var ticketType = sheet.getRange(row, 6).getValue() || "غير معروف";      // Ticket Type
  
  // Get raw display values for dates
  var departureDate = sheet.getRange(row, 7).getDisplayValue();
  var returnDate = sheet.getRange(row, 8).getDisplayValue();
  
  var ticketId = sheet.getRange(row, 9).getValue() || "غير معروف";        // Ticket ID
  var employeeOperations = sheet.getRange(row, 10).getValue() || "غير محدد"; // Employee - Operations
  var employeeSales = sheet.getRange(row, 11).getValue() || "غير محدد";   // Employee - Sales
  var purchaseFrom = sheet.getRange(row, 12).getValue() || "غير محدد";    // Purchase From
  var purchaseValue = sheet.getRange(row, 13).getValue() || "غير محدد";   // Purchase Value
  var soldTo = sheet.getRange(row, 14).getValue() || "غير محدد";          // Sold To
  var soldValue = sheet.getRange(row, 15).getValue() || "غير محدد";       // Sold Value
  var Passport = sheet.getRange(row, 16).getValue() || "غير محدد";           // Passport
  
  // Set title based on whether this is a new ticket or an edit
  var title = type === "edit" ? "🎫 *تعديل تذكرة سابقة*" : "🎫 *تثبيت تذكرة جديدة*";
  
  // Format the message with all ticket details
  return title +
         "\n\n📧 *المٌدخل للبيانات*: " + email + 
         "\n👥 *اسم المسافر(ين)*: " + travelerName + 
         "\n🛫 *المغادرة من*: " + departure + 
         "\n🛬 *الوصول إلى*: " + arrival + 
         "\n🎟️ *نوع التذكرة*: " + ticketType + 
         "\n📅 *تاريخ المغادرة*: " + departureDate + 
         "\n🔄 *تاريخ العودة*: " + returnDate + 
         "\n🆔 *رقم التذكرة*: " + ticketId + 
         "\n👨‍💼 *موظف العمليات*: " + employeeOperations + 
         "\n👨‍💼 *موظف المبيعات*: " + employeeSales + 
         "\n💰 *الشراء من*: " + purchaseFrom + 
         "\n💵 *قيمة الشراء*: " + purchaseValue + 
         "\n👤 *بيعت إلى*: " + soldTo + 
         "\n💰 *قيمة البيع*: " + soldValue + 
         "\n✍🏻 *جواز السفر*: " + Passport;
}

// Helper function to get changed fields for an edited ticket
function getChangedFields(sheet, row, currentValues) {
  var ticketId = sheet.getRange(row, 9).getValue();
  var changedFields = [];
  
  try {
    // Get current display values
    var currentDisplayValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    
    // Get previous values from the MessageIDs sheet
    var previousValues = getPreviousValues(sheet, row, currentDisplayValues);
    
    // Get headers for field names
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    
    // Compare fields and collect changes
    for (var i = 0; i < headers.length; i++) {
      // Skip timestamp, email and edit URL columns
      if (i == 0 || i == 1 || i == 16) continue;
      
      var currentVal = currentDisplayValues[i] || "";
      var prevVal = previousValues[i] || "";
      
      // Handle date comparisons using raw display values
      if (currentVal !== prevVal) {
        changedFields.push(`🔄 *${headers[i]}*: ${prevVal} ➡️ ${currentVal}`);
      }
    }
  } catch (error) {
    Logger.log("Error finding changed fields: " + error.message);
  }
  
  return changedFields;
}

// Helper function to format dates consistently for display
function formatDateForDisplay(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, "GMT+3", "dd/MM/yyyy HH:mm:ss");
  }
  return String(date); // Return raw value if it's already a string
}

// Helper function to pad single digits with zero
function padZero(num) {
  return num < 10 ? "0" + num : num;
}

// ====================================================================
// 📌 وظيفة: إضافة تذكرة (مكانها محجوز فقط)
function addTicket(chatId) {
  // هذه وظيفة مبدئية Placeholder
  sendMessage(chatId, "✅ تمت إضافة التذكرة بنجاح!");
}


// ====================================================================
// 📌 Setup MessageIDs Sheet - This runs once to create the sheet if it doesn't exist
function setupMessageIDsSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var messageIDsSheet = ss.getSheetByName("MessageIDs");
    
    if (!messageIDsSheet) {
      messageIDsSheet = ss.insertSheet("MessageIDs");
      var headers = ["TicketID", "MessageID", "LastEditTimestamp", "PreviousValues"];
      messageIDsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      messageIDsSheet.setFrozenRows(1);
    }
  }
  
  // Helper function to store message ID and previous values
  function storeMessageInfo(ticketId, messageId, values) {
    setupMessageIDsSheet(); // Ensure sheet exists
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("MessageIDs");
    
    // Convert values to JSON for storage
    var valuesJson = JSON.stringify(values);
    var timestamp = new Date();
    
    // Check if ticketId already exists
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === ticketId) {
        // Update existing row
        sheet.getRange(i+1, 2).setValue(messageId);
        sheet.getRange(i+1, 3).setValue(timestamp);
        sheet.getRange(i+1, 4).setValue(valuesJson);
        return;
      }
    }
    
    // If not found, add new row
    sheet.appendRow([ticketId, messageId, timestamp, valuesJson]);
  }
  
  // Helper function to get previous values from the MessageIDs sheet
  function getPreviousValues(sheet, row, currentValues) {
    var ticketId = sheet.getRange(row, 9).getValue();
    
    if (!ticketId) return currentValues;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var messageIDsSheet = ss.getSheetByName("MessageIDs");
    
    // If MessageIDs sheet doesn't exist, create it
    if (!messageIDsSheet) {
      setupMessageIDsSheet();
      return currentValues; // Return current values as there's no history yet
    }
    
    // Look for the ticket ID in the MessageIDs sheet
    var data = messageIDsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === ticketId) {
        try {
          // Parse the stored JSON values
          var previousValues = JSON.parse(data[i][3]);
          return previousValues;
        } catch (e) {
          Logger.log("Error parsing previous values: " + e.toString());
          return currentValues;
        }
      }
    }
    
    // If not found, return current values
    return currentValues;
  }
  
  


// ------------------------------------------------------------------ Analytics Dashboard System

/**
 * Creates an analytics dashboard for the bot
 * @param {string} chatId - The Telegram chat ID to send the dashboard to
 */
function showAnalyticsDashboard(chatId) {
  // Create two columns of buttons for time periods
  var buttons = [
    [
      { text: "📊 هذا اليوم", callback_data: "analytics_period_this_day" },
      { text: "📊 اليوم السابق", callback_data: "analytics_period_last_day" }
    ],
    [
      { text: "📊 هذا الأسبوع", callback_data: "analytics_period_this_week" },
      { text: "📊 الأسبوع السابق", callback_data: "analytics_period_last_week" }
    ],
    [
      { text: "📊 هذا الشهر", callback_data: "analytics_period_this_month" },
      { text: "📊 الشهر السابق", callback_data: "analytics_period_last_month" }
    ],
    [
      { text: "📊 هذا الربع", callback_data: "analytics_period_this_quarter" },
      { text: "📊 الربع السابق", callback_data: "analytics_period_last_quarter" }
    ],
    [
      { text: "📊 هذا العام", callback_data: "analytics_period_this_year" },
      { text: "📊 العام السابق", callback_data: "analytics_period_last_year" }
    ],
    [
      { text: "📊 كل البيانات", callback_data: "analytics_period_all" }
    ],
    [
      { text: "🔙 العودة للقائمة الرئيسية", callback_data: "back_to_main" }
    ]
  ];

  sendMessage(chatId, "📊 اختر فترة التحليل:", { inline_keyboard: buttons });
}

// إضافة دالة لإنشاء نطاق تاريخ بناءً على الفترة المختارة
function getDateRangeForPeriod(period) {
  var now = new Date();
  var startDate, endDate = now;
  
  switch(period) {
    case "this_day":
      startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      break;
    case "last_day":
      startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
      endDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 23, 59, 59);
      break;
    case "this_week":
      // الحصول على أول يوم من الأسبوع الحالي (الأحد)
      var day = now.getDay(); // 0 للأحد، 1 للاثنين، إلخ
      startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - day);
      break;
    case "last_week":
      var day = now.getDay();
      startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - day - 7);
      endDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - day - 1, 23, 59, 59);
      break;
    case "this_month":
      startDate = new Date(now.getFullYear(), now.getMonth(), 1);
      break;
    case "last_month":
      startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      endDate = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
      break;
    case "this_quarter":
      var quarter = Math.floor(now.getMonth() / 3);
      startDate = new Date(now.getFullYear(), quarter * 3, 1);
      break;
    case "last_quarter":
      var quarter = Math.floor(now.getMonth() / 3);
      startDate = new Date(now.getFullYear(), (quarter - 1) * 3, 1);
      if (quarter === 0) {
        startDate = new Date(now.getFullYear() - 1, 9, 1); // Q4 of previous year
      }
      endDate = new Date(now.getFullYear(), quarter * 3, 0, 23, 59, 59);
      break;
    case "this_year":
      startDate = new Date(now.getFullYear(), 0, 1);
      break;
    case "last_year":
      startDate = new Date(now.getFullYear() - 1, 0, 1);
      endDate = new Date(now.getFullYear() - 1, 11, 31, 23, 59, 59);
      break;
    case "all":
    default:
      startDate = new Date(2000, 0, 1); // تاريخ قديم كبداية
      break;
  }
  
  return { startDate: startDate, endDate: endDate };
}

// دالة تحليل التذاكر حسب الفترة المختارة
function analyzeTicketsForPeriod(chatId, period) {
  try {
    // الحصول على نطاق التاريخ المطلوب
    var dateRange = getDateRangeForPeriod(period);
    var startDate = dateRange.startDate;
    var endDate = dateRange.endDate;
    
    // تنسيق التواريخ للعرض
    var formattedStartDate = Utilities.formatDate(startDate, "GMT+3", "yyyy/MM/dd");
    var formattedEndDate = Utilities.formatDate(endDate, "GMT+3", "yyyy/MM/dd");
    
    Logger.log("تحليل التذاكر للفترة من " + formattedStartDate + " إلى " + formattedEndDate);
    
    // إرسال رسالة للمستخدم لإظهار التقدم
    sendMessage(chatId, "⏳ جاري تحليل البيانات للفترة من " + formattedStartDate + " إلى " + formattedEndDate + "...");
    
    // الحصول على بيانات التذاكر
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
    var data = sheet.getDataRange().getValues();
    var headers = data[0]; // صف العناوين
    var rows = data.slice(1); // تخطي صف العناوين
    
    Logger.log("عدد الصفوف الكلي: " + rows.length);
    
    // استخراج فهارس الأعمدة من العناوين
    var TICKET_COLUMNS = {
      TIMESTAMP: -1,
      EMAIL: -1,
      TRAVELER_NAME: -1,
      DEPARTURE: -1,
      ARRIVAL: -1,
      TICKET_TYPE: -1,
      DEPARTURE_DATE: -1,
      RETURN_DATE: -1,
      TICKET_ID: -1,
      EMPLOYEE_OPS: -1,
      EMPLOYEE_SALES: -1,
      PURCHASE_FROM: -1,
      PURCHASE_VALUE: -1,
      SOLD_TO: -1,
      SOLD_VALUE: -1,
      Passport: -1,
      EDIT: -1,
      STATUS: -1
    };
    
    // العثور على فهارس الأعمدة من أسماء العناوين
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i]).trim();
      if (header.includes("Timestamp")) TICKET_COLUMNS.TIMESTAMP = i;
      else if (header.includes("Email")) TICKET_COLUMNS.EMAIL = i;
      else if (header.includes("Traveler")) TICKET_COLUMNS.TRAVELER_NAME = i;
      else if (header.includes("Departure") && !header.includes("Date")) TICKET_COLUMNS.DEPARTURE = i;
      else if (header.includes("Arrival")) TICKET_COLUMNS.ARRIVAL = i;
      else if (header.includes("Ticket Type")) TICKET_COLUMNS.TICKET_TYPE = i;
      else if (header.includes("Departure Date")) TICKET_COLUMNS.DEPARTURE_DATE = i;
      else if (header.includes("Return Date")) TICKET_COLUMNS.RETURN_DATE = i;
      else if (header.includes("Ticket ID")) TICKET_COLUMNS.TICKET_ID = i;
      else if (header.includes("Employee") && header.includes("Operations")) TICKET_COLUMNS.EMPLOYEE_OPS = i;
      else if (header.includes("Employee") && header.includes("Sales")) TICKET_COLUMNS.EMPLOYEE_SALES = i;
      else if (header.includes("Purchase From")) TICKET_COLUMNS.PURCHASE_FROM = i;
      else if (header.includes("Purchase Value")) TICKET_COLUMNS.PURCHASE_VALUE = i;
      else if (header.includes("Sold To")) TICKET_COLUMNS.SOLD_TO = i;
      else if (header.includes("Sold Value")) TICKET_COLUMNS.SOLD_VALUE = i;
      else if (header.includes("Passport")) TICKET_COLUMNS.Passport = i;
      else if (header.includes("Edit")) TICKET_COLUMNS.EDIT = i;
      else if (header.includes("Status")) TICKET_COLUMNS.STATUS = i;
    }
    
    // التحقق من أن الأعمدة المطلوبة موجودة
    if (TICKET_COLUMNS.TIMESTAMP === -1 || TICKET_COLUMNS.STATUS === -1) {
      sendMessage(chatId, "❌ لم يتم العثور على أعمدة مطلوبة في جدول البيانات. الرجاء التحقق من تنسيق الجدول.");
      return;
    }
    
    // تصفية التذاكر حسب التاريخ
    var filteredTickets = rows.filter(function(row) {
      // استخدام حقل Timestamp للتصفية
      var timestampStr = row[TICKET_COLUMNS.TIMESTAMP];
      var timestamp;
      
      if (timestampStr instanceof Date) {
        timestamp = timestampStr;
      } else if (typeof timestampStr === 'string') {
        // تحويل من نص إلى تاريخ
        // محاولة تحليل التاريخ بعدة تنسيقات
        timestamp = parseDate(timestampStr);
      }
      
      if (!timestamp || isNaN(timestamp.getTime())) {
        return false; // تخطي الصفوف بدون تاريخ صالح
      }
      
      // التحقق مما إذا كان التاريخ ضمن النطاق المطلوب
      return timestamp >= startDate && timestamp <= endDate;
    });
    
    Logger.log("عدد التذاكر بعد التصفية: " + filteredTickets.length);
    
    // إذا لم تكن هناك تذاكر
    if (filteredTickets.length === 0) {
      sendMessage(chatId, `❌ لا توجد تذاكر في هذه الفترة.`);
      return;
    }
    
    // حساب الإحصائيات
    var totalTickets = filteredTickets.length;
    var openTickets = filteredTickets.filter(row => String(row[TICKET_COLUMNS.STATUS] || "").trim() === "مفتوحة").length;
    var closedTickets = filteredTickets.filter(row => String(row[TICKET_COLUMNS.STATUS] || "").trim() === "مغلقة").length;
    
    // حساب إجمالي قيم البيع والشراء
    var totalSoldValue = filteredTickets.reduce((sum, row) => {
      var value = parseFloat(row[TICKET_COLUMNS.SOLD_VALUE]);
      return sum + (isNaN(value) ? 0 : value);
    }, 0);
    
    var totalPurchaseValue = filteredTickets.reduce((sum, row) => {
      var value = parseFloat(row[TICKET_COLUMNS.PURCHASE_VALUE]);
      return sum + (isNaN(value) ? 0 : value);
    }, 0);
    
    var profit = totalSoldValue - totalPurchaseValue;
    
    // تنسيق الأرقام
    var formattedSoldValue = totalSoldValue.toLocaleString('ar-SA');
    var formattedPurchaseValue = totalPurchaseValue.toLocaleString('ar-SA');
    var formattedProfit = profit.toLocaleString('ar-SA');
    
    // تحليل حسب نوع التذكرة (one-way vs round-trip)
    var ticketTypes = {};
    filteredTickets.forEach(function(row) {
      var type = String(row[TICKET_COLUMNS.TICKET_TYPE] || "").trim();
      if (!type) return;
      
      if (!ticketTypes[type]) {
        ticketTypes[type] = { total: 0, open: 0, closed: 0 };
      }
      ticketTypes[type].total++;
      
      var status = String(row[TICKET_COLUMNS.STATUS] || "").trim();
      if (status === "مفتوحة") {
        ticketTypes[type].open++;
      } else if (status === "مغلقة") {
        ticketTypes[type].closed++;
      }
    });
    
    // ترتيب أنواع التذاكر حسب العدد
    var ticketTypeStats = Object.entries(ticketTypes)
      .sort((a, b) => b[1].total - a[1].total);
    
    // إنشاء رسالة الإحصائيات
    var message = `📊 <b>إحصائيات التذاكر</b>\n`;
    message += `📝 <b>إجمالي التذاكر:</b> ${totalTickets}\n`;
    message += `✅ <b>التذاكر المغلقة:</b> ${closedTickets} (${Math.round(closedTickets/totalTickets*100 || 0)}%)\n`;
    message += `⏳ <b>التذاكر المفتوحة:</b> ${openTickets} (${Math.round(openTickets/totalTickets*100 || 0)}%)\n\n`;
    
    // إضافة معلومات المبيعات والأرباح
    message += `💰 <b>إجمالي المبيعات:</b> ${formattedSoldValue}\n`;
    message += `💼 <b>إجمالي المشتريات:</b> ${formattedPurchaseValue}\n`;
    message += `📈 <b>الربح الإجمالي:</b> ${formattedProfit}\n\n`;
    
    // إضافة معلومات أنواع التذاكر
    if (ticketTypeStats.length > 0) {
      message += `👨‍💼 <b>أداء الموظفين:</b>\n`;
      ticketTypeStats.forEach((type, index) => {
        var stats = type[1];
        var closingRate = Math.round(stats.closed / stats.total * 100 || 0);
        message += `${index + 1}. ${type[0]}: ${stats.total} تذكرة (${closingRate}% مغلقة)\n`;
      });
    }
    
    // إرسال التحليل للمستخدم
    var inlineKeyboard = {
      inline_keyboard: [
        [
          { text: "👨‍💼 تحليل الموظفين", callback_data: "analytics_employees_" + period },
          { text: "📊 تصدير البيانات", callback_data: "analytics_export_" + period }
        ],
        [{ text: "🔙 الرجوع للتحليلات", callback_data: "show_analytics" }]
      ]
    };
    
    Logger.log("DEBUG: Sending analytics results with buttons: " + JSON.stringify(inlineKeyboard));
    sendMessage(chatId, message, inlineKeyboard);
  } catch (error) {
    Logger.log("خطأ في تحليل التذاكر: " + error.message);
    Logger.log(error.stack);
    sendMessage(chatId, "❌ حدث خطأ أثناء تحليل البيانات: " + error.message);
  }
}

// دالة مساعدة لتحليل التاريخ من النصوص المختلفة
function parseDate(dateStr) {
  if (!dateStr) return null;
  
  // If input is already a Date object, return it directly
  if (dateStr instanceof Date) {
    return dateStr;
  }
  
  // If string is already in database format (dd/MM/yyyy HH:mm:ss), parse it directly
  if (typeof dateStr === 'string') {
    var dbFormat = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/);
    if (dbFormat) {
      // Create date without timezone conversion (use UTC to avoid local timezone)
      var d = new Date(Date.UTC(
        parseInt(dbFormat[3]), // year
        parseInt(dbFormat[2]) - 1, // month (0-based)
        parseInt(dbFormat[1]), // day
        parseInt(dbFormat[4]), // hour
        parseInt(dbFormat[5]), // minute
        parseInt(dbFormat[6])  // second
      ));
      return d;
    }
  }
  
  // محاولة تحليل بعدة تنسيقات شائعة
  var formats = [
    // dd/MM/yyyy HH:mm:ss without UTC adjustment
    function(s) {
      var parts = s.match(/(\d+)\/(\d+)\/(\d+)\s+(\d+):(\d+):(\d+)/);
      if (parts) {
        return new Date(Date.UTC(
          parseInt(parts[3]), // year
          parseInt(parts[2]) - 1, // month (0-based)
          parseInt(parts[1]), // day
          parseInt(parts[4]), // hour
          parseInt(parts[5]), // minute
          parseInt(parts[6])  // second
        ));
      }
      return null;
    },
    // dd/MM/yyyy
    function(s) {
      var parts = s.match(/(\d+)\/(\d+)\/(\d+)/);
      if (parts) {
        return new Date(Date.UTC(parts[3], parts[2]-1, parts[1]));
      }
      return null;
    },
    // yyyy-MM-dd
    function(s) {
      var parts = s.match(/(\d+)-(\d+)-(\d+)/);
      if (parts) {
        return new Date(Date.UTC(parts[1], parts[2]-1, parts[3]));
      }
      return null;
    }
  ];
  
  for (var i = 0; i < formats.length; i++) {
    var date = formats[i](dateStr);
    if (date && !isNaN(date.getTime())) {
      return date;
    }
  }
  
  // إذا لم تنجح أي طريقة، نجرب بناء كائن تاريخ مباشرة
  var date = new Date(dateStr);
  if (!isNaN(date.getTime())) {
    return date;
  }
  
  return null;
}

/**
 * Creates an Excel file with analytics data in a specific folder and sends a download link
 * @param {string} chatId - The Telegram chat ID to send the export to
 * @param {string} period - The time period to filter data by
 */
function exportAnalyticsToExcel(chatId, period) {
  sendMessage(chatId, "⚙️ جاري إنشاء ملف التقرير...");
  
  var mainSS = SpreadsheetApp.getActiveSpreadsheet();
  var ticketSheet = mainSS.getSheetByName(sheetTickets);
  var data = ticketSheet.getDataRange().getValues();
  var headers = data[0]; // صف العناوين
  var rows = data.slice(1); // تخطي صف العناوين
  
  // استخراج فهارس الأعمدة من العناوين
  var TICKET_COLUMNS = {
    TIMESTAMP: -1,
    TICKET_ID: -1,
    TRAVELER_NAME: -1,
    DEPARTURE_LOCATION: -1,
    STATUS: -1,
    EMPLOYEE_SALES: -1,
    EMPLOYEE_OPERATIONS: -1
  };
  
  // العثور على فهارس الأعمدة من أسماء العناوين
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i]).trim();
    if (header.includes("Timestamp")) TICKET_COLUMNS.TIMESTAMP = i;
    else if (header.includes("Ticket ID")) TICKET_COLUMNS.TICKET_ID = i;
    else if (header.includes("Traveler")) TICKET_COLUMNS.TRAVELER_NAME = i;
    else if (header.includes("Departure") && !header.includes("Date")) TICKET_COLUMNS.DEPARTURE_LOCATION = i;
    else if (header.includes("Status")) TICKET_COLUMNS.STATUS = i;
    else if (header.includes("Employee") && header.includes("Sales")) TICKET_COLUMNS.EMPLOYEE_SALES = i;
    else if (header.includes("Employee") && header.includes("Operations")) TICKET_COLUMNS.EMPLOYEE_OPERATIONS = i;
  }
  
  // التحقق من أن الأعمدة المطلوبة موجودة
  if (TICKET_COLUMNS.TIMESTAMP === -1 || TICKET_COLUMNS.TRAVELER_NAME === -1 || TICKET_COLUMNS.STATUS === -1) {
    sendMessage(chatId, "❌ لم يتم العثور على أعمدة مطلوبة في جدول البيانات. الرجاء التحقق من تنسيق الجدول.");
    return;
  }
  
  // تصفية البيانات حسب الفترة إذا تم تحديدها
  var filteredRows = rows;
  var periodStr = "كل_البيانات";
  
  if (period) {
    // الحصول على نطاق التاريخ المطلوب
    var dateRange = getDateRangeForPeriod(period);
    var startDate = dateRange.startDate;
    var endDate = dateRange.endDate;
    
    // تحديث اسم التقرير ليعكس الفترة
    periodStr = Utilities.formatDate(startDate, "GMT+3", "yyyy-MM-dd") + "_to_" + 
                Utilities.formatDate(endDate, "GMT+3", "yyyy-MM-dd");
    
    // تصفية الصفوف حسب التاريخ
    filteredRows = rows.filter(function(row) {
      // استخدام حقل Timestamp للتصفية
      var timestampStr = row[TICKET_COLUMNS.TIMESTAMP];
      var timestamp;
      
      if (timestampStr instanceof Date) {
        timestamp = timestampStr;
      } else if (typeof timestampStr === 'string') {
        timestamp = parseDate(timestampStr);
      }
      
      if (!timestamp || isNaN(timestamp.getTime())) {
        return false; // تخطي الصفوف بدون تاريخ صالح
      }
      
      // التحقق مما إذا كان التاريخ ضمن النطاق المطلوب
      return timestamp >= startDate && timestamp <= endDate;
    });
  }
  
  // Create a new spreadsheet
  var reportName = "تقرير_" + periodStr + "_" + Utilities.formatDate(new Date(), "GMT+3", "yyyy_MM_dd");
  var newSS = SpreadsheetApp.create(reportName);
  
  // Move to specified folder
  try {
    var folder = DriveApp.getFolderById('1Sx9Yo3DkEtCbgd6kpvbDKQ6P27PESyH9');
    var file = DriveApp.getFileById(newSS.getId());
    folder.addFile(file);
    // Remove from root folder
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    sendMessage(chatId, "❌ خطأ في إنشاء الملف: " + e.message);
    return;
  }
  
  // Prepare the report
  var reportSheet = newSS.getSheets()[0];
  reportSheet.setName("تقرير التذاكر");
  
  // Find the Edit column index
  var editColumnIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().includes("Edit")) {
      editColumnIndex = i;
      break;
    }
  }
  
  // Create new headers and data arrays without the Edit column
  var reportHeaders = [];
  for (var i = 0; i < headers.length; i++) {
    if (i !== editColumnIndex) {
      reportHeaders.push(headers[i]);
    }
  }
  
  var reportData = [];
  for (var i = 0; i < filteredRows.length; i++) {
    var row = filteredRows[i];
    var newRow = [];
    for (var j = 0; j < row.length; j++) {
      if (j !== editColumnIndex) {
        newRow.push(row[j]);
      }
    }
    reportData.push(newRow);
  }
  
  // Add headers without Edit column
  reportSheet.getRange(1, 1, 1, reportHeaders.length)
    .setValues([reportHeaders])
    .setFontWeight("bold");
  
  // Add data without Edit column
  if (reportData.length > 0) {
    reportSheet.getRange(2, 1, reportData.length, reportHeaders.length)
      .setValues(reportData);
  } else {
    // إذا لم تكن هناك بيانات في النطاق المحدد
    reportSheet.getRange(2, 1)
      .setValue("لا توجد بيانات في هذه الفترة");
  }
  
  // Formatting
  reportSheet.autoResizeColumns(1, reportHeaders.length);
  reportSheet.getRange(1, 1, 1, reportHeaders.length)
    .setBackground("#f0f0f0")
    .setFontSize(12);
  
  // Generate shareable link
  var url = newSS.getUrl();
  
  // إضافة معلومات الفترة للرسالة
  var periodMessage = "";
  if (period) {
    var formattedStartDate = Utilities.formatDate(startDate, "GMT+3", "yyyy/MM/dd");
    var formattedEndDate = Utilities.formatDate(endDate, "GMT+3", "yyyy/MM/dd");
    periodMessage = `\nالفترة: ${formattedStartDate} إلى ${formattedEndDate}`;
  }
  
  sendMessage(chatId, `✅ تم إنشاء التقرير بنجاح!${periodMessage}\nعدد السجلات: ${filteredRows.length}\n\nرابط الملف:\n${url}`);
}

/**
 * Shows employee performance analysis
 * @param {string} chatId
 * @param {string} period - The time period to filter data by
 */
function showEmployeeAnalysis(chatId, period) {
  // Check if user is admin
  if (!isAdmin(chatId)) {
    sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم الوصول إلى تحليل أداء الموظفين.");
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getValues();
  var headers = data[0]; // صف العناوين
  var rows = data.slice(1); // تخطي صف العناوين
  
  // استخراج فهارس الأعمدة من العناوين
  var TICKET_COLUMNS = {
    TIMESTAMP: -1,
    EMPLOYEE_OPERATIONS: -1,
    EMPLOYEE_SALES: -1,
    STATUS: -1
  };
  
  // العثور على فهارس الأعمدة من أسماء العناوين
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i]).trim();
    if (header.includes("Timestamp")) TICKET_COLUMNS.TIMESTAMP = i;
    else if (header.includes("Employee") && header.includes("Operations")) TICKET_COLUMNS.EMPLOYEE_OPERATIONS = i;
    else if (header.includes("Employee") && header.includes("Sales")) TICKET_COLUMNS.EMPLOYEE_SALES = i;
    else if (header.includes("Status")) TICKET_COLUMNS.STATUS = i;
  }
  
  // التحقق من أن الأعمدة المطلوبة موجودة
  if (TICKET_COLUMNS.EMPLOYEE_OPERATIONS === -1 || TICKET_COLUMNS.EMPLOYEE_SALES === -1 || TICKET_COLUMNS.STATUS === -1 || TICKET_COLUMNS.TIMESTAMP === -1) {
    sendMessage(chatId, "❌ لم يتم العثور على أعمدة الموظفين أو الحالة في جدول البيانات. الرجاء التحقق من تنسيق الجدول.");
    return;
  }
  
  // تصفية البيانات حسب الفترة إذا تم تحديدها
  var filteredRows = rows;
  if (period) {
    // الحصول على نطاق التاريخ المطلوب
    var dateRange = getDateRangeForPeriod(period);
    var startDate = dateRange.startDate;
    var endDate = dateRange.endDate;
    
    // تصفية الصفوف حسب التاريخ
    filteredRows = rows.filter(function(row) {
      // استخدام حقل Timestamp للتصفية
      var timestampStr = row[TICKET_COLUMNS.TIMESTAMP];
      var timestamp;
      
      if (timestampStr instanceof Date) {
        timestamp = timestampStr;
      } else if (typeof timestampStr === 'string') {
        timestamp = parseDate(timestampStr);
      }
      
      if (!timestamp || isNaN(timestamp.getTime())) {
        return false; // تخطي الصفوف بدون تاريخ صالح
      }
      
      // التحقق مما إذا كان التاريخ ضمن النطاق المطلوب
      return timestamp >= startDate && timestamp <= endDate;
    });
  }
  
  // Sales employee stats
  var salesStats = {};
  filteredRows.forEach(row => {
    var employee = row[TICKET_COLUMNS.EMPLOYEE_SALES] || "غير محدد";
    if (typeof employee !== 'string') employee = String(employee);
    employee = employee.trim();
    
    if (!salesStats[employee]) salesStats[employee] = { total: 0, open: 0, closed: 0 };
    salesStats[employee].total++;
    
    var status = String(row[TICKET_COLUMNS.STATUS] || "").trim();
    if (status === "مفتوحة") salesStats[employee].open++;
    else if (status === "مغلقة") salesStats[employee].closed++;
  });
  
  // Operations employee stats
  var opsStats = {};
  filteredRows.forEach(row => {
    var employee = row[TICKET_COLUMNS.EMPLOYEE_OPERATIONS] || "غير محدد";
    if (typeof employee !== 'string') employee = String(employee);
    employee = employee.trim();
    
    if (!opsStats[employee]) opsStats[employee] = { total: 0, open: 0, closed: 0 };
    opsStats[employee].total++;
    
    var status = String(row[TICKET_COLUMNS.STATUS] || "").trim();
    if (status === "مفتوحة") opsStats[employee].open++;
    else if (status === "مغلقة") opsStats[employee].closed++;
  });
  
  // Build message
  var message = "<b>👨‍💼 تحليل أداء الموظفين</b>\n\n";
  
  // Sales employees section
  message += "<b>📊 موظفي المبيعات:</b>\n";
  Object.keys(salesStats)
    .filter(e => e !== "غير محدد")
    .sort((a, b) => salesStats[b].total - salesStats[a].total)
    .slice(0, 5) // Top 5 employees
    .forEach(employee => {
      var closureRate = Math.round((salesStats[employee].closed / salesStats[employee].total) * 100);
      message += `- ${employee}: ${salesStats[employee].total} تذكرة | معدل الإغلاق: ${closureRate}%\n`;
    });
  message += "\n";
  
  // Operations employees section
  message += "<b>🔧 موظفي العمليات:</b>\n";
  Object.keys(opsStats)
    .filter(e => e !== "غير محدد")
    .sort((a, b) => opsStats[b].total - opsStats[a].total)
    .slice(0, 5) // Top 5 employees
    .forEach(employee => {
      var closureRate = Math.round((opsStats[employee].closed / opsStats[employee].total) * 100);
      message += `- ${employee}: ${opsStats[employee].total} تذكرة | معدل الإغلاق: ${closureRate}%\n`;
    });
  
  var buttons = {
    inline_keyboard: [
      [{ text: "🔙 العودة للإحصائيات", callback_data: "show_analytics" }]
    ]
  };
  
  sendMessage(chatId, message, buttons);
}


// ------------------------------------------------------------------ Ticket Management System

// عرض خيارات الأشهر
function showMonthSelection(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getValues();
  
  // تصفية التذاكر المفتوحة فقط (العمود STATUS يحتوي على قيمة "مفتوحة")
  var openTickets = data.slice(1).filter(row => row[TICKET_COLUMNS.STATUS] === "مفتوحة");

  var monthsSet = {};
  openTickets.forEach(row => {
    var dateStr = row[TICKET_COLUMNS.DEPARTURE_DATE];
    var date = dateStr instanceof Date ? dateStr : new Date(dateStr);
    if (isNaN(date.getTime())) return;

    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    var key = `${year}-${month.toString().padStart(2, '0')}`;
    var display = `${getMonthName(month)} ${year}`;
    monthsSet[key] = display;
  });

  if (!Object.keys(monthsSet).length) {
    sendMessage(chatId, "لا توجد تذاكر مفتوحة في أي شهر.");
    return;
  }

  var buttons = Object.entries(monthsSet).map(([key, display]) => [{
    text: display,
    callback_data: `month_${key}`
  }]);

  sendMessage(chatId, "اختر الشهر لعرض التذاكر المفتوحة:", { inline_keyboard: buttons });
}

// عرض اسماء الأشهر
function getMonthName(monthNumber) {
  var months = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
  return months[monthNumber - 1];
}

// عرض التكتات حسب الشهر المختار
function showTicketsForMonth(chatId, monthKey) {
  var [year, month] = monthKey.split('-').map(Number);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getValues();

  var filteredTickets = data.filter(row => {
    if (row[TICKET_COLUMNS.STATUS] !== "مفتوحة") return false;

    var date = row[TICKET_COLUMNS.DEPARTURE_DATE] instanceof Date ? 
               row[TICKET_COLUMNS.DEPARTURE_DATE] : 
               new Date(row[TICKET_COLUMNS.DEPARTURE_DATE]);
               
    if (isNaN(date.getTime())) return false;

    return date.getFullYear() === year && date.getMonth() + 1 === month;
  });

  if (!filteredTickets.length) {
    sendMessage(chatId, `لا توجد تذاكر مفتوحة في ${getMonthName(month)} ${year}.`);
    return;
  }

var buttons = filteredTickets.map(row => {
  const departure = row[TICKET_COLUMNS.DEPARTURE_LOCATION];
  const arrival = row[TICKET_COLUMNS.ARRIVAL_LOCATION];
  const ticketId = row[TICKET_COLUMNS.TICKET_ID];

  // إضافة علامات التحكم لفرض اتجاه النص
  const rtlMark = "\u200F"; // Right-to-Left Mark (للنصوص العربية)
  const ltrMark = "\u200E"; // Left-to-Right Mark (للأرقام/الإنجليزية)
  
  return [{
    text: 
      `${rtlMark}${departure} - ${arrival}${rtlMark}\n` + 
      `${ltrMark}#${ticketId}`,
    callback_data: `ticket_${ticketId}_${monthKey}`
  }];
});

sendMessage(chatId, `📂 التذاكر المفتوحة في ${getMonthName(month)} ${year}:`, { inline_keyboard: buttons });
}





// التعديل على دالة startTicketConversation
function startTicketConversation(chatId) {
  Logger.log("Opening ticket link for chatId: " + chatId);

  // إعداد الزر لفتح رابط البوت الثاني
  var replyMarkup = {
    inline_keyboard: [
      [
        {
          text: "➕ إضافة تذكرة",
          url: "https://t.me/Tickets321_bot/AddTickets"  // رابط البوت الثاني
        }
      ]
    ]
  };

  // إرسال رسالة مع الزر
  sendMessage(chatId, "🌐 اضغط الزر لإضافة تذكرة  :", replyMarkup);
}

// عرض التذاكر المفتوحة
function showTickets(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getValues();

  // تصفية التذاكر بحيث تظهر فقط تلك التي حالتها "مفتوحة"
  var filteredTickets = data.filter(row => row[5] === "مفتوحة");

  if (filteredTickets.length === 0) {
    sendMessage(chatId, "😞 لا توجد تذاكر مفتوحة حالياً.");
    return;
  }

  // إنشاء أزرار للقائمة باستخدام البيانات المفلترة
  var buttons = filteredTickets.map(row => [{
    text: `🔖 #${row[1]} - ${row[3]}`,  // استخدم علامات الاقتباس بدلاً من النص الغير صحيح
    callback_data: "ticket_" + row[1]
  }]);

  sendMessage(chatId, "📂 التذاكر المفتوحة:", { inline_keyboard: buttons });
}

// عرض تفاصيل التذكرة مع زر إغلاق وزر عودة
function showTicketDetails(chatId, ticketId, monthKey) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getDisplayValues(); // Use display values
  var ticket = data.find(row => String(row[TICKET_COLUMNS.TICKET_ID]) == String(ticketId));

  if (!ticket) {
    sendMessage(chatId, "❌ التذكرة غير موجودة.");
    return;
  }
  
  // Handle departure date
  var departureDate = ticket[TICKET_COLUMNS.DEPARTURE_DATE];
  var formattedDepartureDate = departureDate;
  
  // Handle return date
  var returnDate = ticket[TICKET_COLUMNS.RETURN_DATE];
  var returnDateText = "";
  
  if (returnDate) {
    formattedReturnDate = returnDate;
    returnDateText = `\n📅 <b>تاريخ العودة:</b> ${formattedReturnDate}\n\n`;
  }

  var message = 
    `🎫 <b>تذكرة رقم:</b> #${ticket[TICKET_COLUMNS.TICKET_ID]}\n` +
    `👤 <b>اسم المسافر:</b> ${ticket[TICKET_COLUMNS.TRAVELER_NAME]}\n` +
    `✈️ <b>الرحلة:</b> من ${ticket[TICKET_COLUMNS.DEPARTURE_LOCATION]}\n  → إلى ${ticket[TICKET_COLUMNS.ARRIVAL_LOCATION]}\n\n` +
    `🎫 <b>نوع التذكرة:</b> ${ticket[TICKET_COLUMNS.TICKET_TYPE]}\n\n` +
    `📅 <b>تاريخ المغادرة:</b> ${formattedDepartureDate}${returnDateText}\n\n` +
    `👨‍💼 <b>موظف المبيعات:</b> ${ticket[TICKET_COLUMNS.EMPLOYEE_SALES]}\n` +
    `👨‍💻 <b>موظف العمليات:</b> ${ticket[TICKET_COLUMNS.EMPLOYEE_OPERATIONS]}\n\n` +
    `💲 <b>تم الشراء من:</b> ${ticket[TICKET_COLUMNS.PURCHASE_FROM]} (${ticket[TICKET_COLUMNS.PURCHASE_VALUE]})\n` +
    `💰 <b>تم البيع لـ:</b> ${ticket[TICKET_COLUMNS.SOLD_TO]} (${ticket[TICKET_COLUMNS.SOLD_VALUE]})\n\n` +
    `📝 <b>جواز السفر:</b> ${ticket[TICKET_COLUMNS.Passport]}\n` +
    `🚦 <b>الحالة:</b> ${ticket[TICKET_COLUMNS.STATUS]}`;

  var replyMarkup = {
    inline_keyboard: [
      [
        { text: "❌ إغلاق التذكرة", callback_data: `close_ticket_${monthKey}_${ticketId}` },
        { text: "🔄 تبديل الحالة", callback_data: `toggle_ticket_status_${ticketId}_${monthKey}` }
      ],
      [
        { text: "✏️ تعديل التذكرة", callback_data: `edit_ticket_${ticketId}_${monthKey}` },
        { text: "🔙 العودة للقائمة", callback_data: `back_to_month_${monthKey}` }
      ]
    ]
  };

  sendMessage(chatId, message, replyMarkup);
}

// وظيفة مساعدة لتنسيق التواريخ
function formatDate(dateStr) {
  return dateStr ? String(dateStr) : "غير محدد";
}

// إغلاق التذكرة
function closeTicket(chatId, ticketId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
    var data = sheet.getDataRange().getValues();
    var ticketFound = false;
    
    // Initialize ticket columns if not already initialized
    if (TICKET_COLUMNS.STATUS === null) {
      initializeTicketColumns();
    }
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][TICKET_COLUMNS.TICKET_ID]) === String(ticketId)) {
        // تحديث حالة التذكرة إلى "مغلقة"
        sheet.getRange(i + 1, TICKET_COLUMNS.STATUS + 1).setValue("مغلقة");
        ticketFound = true;
        sendMessage(chatId, "✅ تم إغلاق التذكرة بنجاح!");
        return true;
      }
    }
    
    if (!ticketFound) {
      sendMessage(chatId, "❌ لم يتم العثور على التذكرة رقم " + ticketId);
      return false;
    }
  } catch (error) {
    Logger.log("Error in closeTicket: " + error.message);
    sendMessage(chatId, "❌ حدث خطأ أثناء محاولة إغلاق التذكرة: " + error.message);
    return false;
  }
  
  return false;
}


function editTicket(chatId, ticketId) {
  if (!isAdmin(chatId)) {
    sendMessage(chatId, "⛔️ فقط المشرفين يمكنهم تعديل التذاكر.");
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTickets);
  var data = sheet.getDataRange().getValues();
  var ticket = data.find(row => row[1] == ticketId);

  if (!ticket) {
    sendMessage(chatId, "❌ لم يتم العثور على التذكرة.");
    return;
  }

  var editLink = ticket[12]; // العمود M (مؤشر 12 لأن المؤشر يبدأ من 0)

  if (!editLink || editLink.trim() === "") {
    sendMessage(chatId, "⚠️ لا يوجد رابط تعديل لهذه التذكرة.");
    return;
  }

  // إرسال الرابط بصيغة HTML قابلة للنقر
  sendMessage(chatId, `🔗 <b>رابط تعديل التذكرة:</b>\n<a href="${editLink}">${editLink}</a>`, { parse_mode: "HTML" });
}


// ------------------------------------------------------------------ User Management System

// Show the main user management menu with options
function showUserManagementMenu(chatId) {
  // Check if user is an admin
  if (!isAdmin(chatId)) {
    sendMessage(chatId, "⛔ عذراً، هذه الخاصية متاحة للمشرفين فقط.");
    return;
  }
  
  var keyboard = {
    inline_keyboard: [
      [{ text: "👥 إدارة المستخدمين الأساسيين", callback_data: "user_manage_main" }],
      [{ text: "📢 إدارة قائمة البث", callback_data: "user_manage_broadcast" }],
      [{ text: "🔙 العودة للقائمة الرئيسية", callback_data: "back_to_main" }]
    ]
  };
  
  sendMessage(chatId, "🛠️ <b>نظام إدارة المستخدمين</b>\n\nاختر إحدى الخيارات التالية:", keyboard);
}

// Show management options for main users
function showMainUsersManagement(chatId) {
  var keyboard = {
    inline_keyboard: [
      [
        {
          text: "➕ إضافة مستخدم جديد",
          url: "https://t.me/Tickets321_bot/user1mangmaent"
        }
      ],
      [{ text: "📋 عرض المستخدمين", callback_data: "list_main_users" }],
      [{ text: "🔙 عودة", callback_data: "user_management_main" }]
    ]
  };

  sendMessage(chatId, "👥 <b>إدارة المستخدمين الأساسيين</b>\n\nاختر إحدى العمليات التالية:", keyboard);
}

// Show management options for broadcast list
function showBroadcastUsersManagement(chatId) {
  var keyboard = {
    inline_keyboard: [
      [{
        text: "➕ إضافة مستخدم للبث",
        url: "https://t.me/Tickets321_bot/users2" // رابط البوت أو الصفحة المطلوبة
      }],
      [{ text: "📋 عرض قائمة البث", callback_data: "list_broadcast_users" }],
      [{ text: "🔙 عودة", callback_data: "user_management_main" }]
    ]
  };

  sendMessage(chatId, "📢 <b>إدارة قائمة البث</b>\n\nاختر إحدى العمليات التالية:", keyboard);
}


// List all main users
function listMainUsers(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers1);
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    sendMessage(chatId, "📝 لا يوجد مستخدمين مسجلين بعد.");
    return;
  }
  
  var message = "👥 <b>قائمة المستخدمين الأساسيين:</b>\n\n";
  var keyboard = {
    inline_keyboard: []
  };
  
  // Skip header row
  for (var i = 1; i < data.length; i++) {
    var userId = data[i][1];
    var permission = data[i][3] || "مستخدم عادي";
    var name = data[i][2] || "بدون اسم";
    
    message += i + ". " + name + " (" + userId + ") - " + permission + "\n";
    keyboard.inline_keyboard.push([
      { text: "🔃 " + name, callback_data: "edit_main_user_" + userId }
    ]);
  }
  
  keyboard.inline_keyboard.push([
    { text: "🔙 عودة", callback_data: "user_manage_main" }
  ]);
  
  sendMessage(chatId, message, keyboard);
}

// List all broadcast users
function listBroadcastUsers(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users2");
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) { // فقط العنوان موجود أو لا يوجد بيانات
    sendMessage(chatId, "📝 لا يوجد مستخدمين في قائمة البث بعد.");
      return;
    }
  
  var message = "📢 <b>قائمة مستخدمي البث:</b>\n\n";
  var keyboard = {
    inline_keyboard: []
  };
  
  for (var i = 1; i < data.length; i++) { // نبدأ من 1 لتجاهل العنوان
    var userId = data[i][1];
    var name = data[i][2] || "بدون اسم";
    
    message += (i) + ". " + name + " (" + userId + ")\n"; // i هو الترتيب الصحيح الآن
    keyboard.inline_keyboard.push([
      { text: "❌ " + name, callback_data: "delete_broadcast_user_" + userId }
    ]);
  }
  
  keyboard.inline_keyboard.push([
    { text: "🔙 عودة", callback_data: "user_manage_broadcast" }
  ]);
  
  sendMessage(chatId, message, keyboard);
}

// Start the process to add a new main user
function startAddMainUser(chatId) {
  userSessionManager.updateContext(chatId, { waitingFor: 'add_main_user_id' });
  sendMessage(chatId, "👤 الرجاء إدخال معرف المستخدم (user ID):");
}

// Start the process to add a new broadcast user
function startAddBroadcastUser(chatId) {
  userSessionManager.updateContext(chatId, { waitingFor: 'add_broadcast_user_id' });
  sendMessage(chatId, "👤 الرجاء إدخال معرف مستخدم البث (user ID):");
}

// Process adding a new main user - step 1 (chatId)
function processAddMainUserStep1(chatId, text) {
  // Validate that the input is a valid chatId (number)
  var userId = text.trim();
  if (isNaN(userId)) {
    sendMessage(chatId, "⚠️ معرف المستخدم يجب أن يكون رقمًا. الرجاء المحاولة مرة أخرى:");
    return;
  }
  
  userSessionManager.updateContext(chatId, { 
    waitingFor: 'add_main_user_name',
    userId: userId
  });
  
  sendMessage(chatId, "👤 الرجاء إرسال اسم المستخدم:");
}

// Process adding a new main user - step 2 (name)
function processAddMainUserStep2(chatId, text) {
  var name = text.trim();
  if (!name) {
    sendMessage(chatId, "⚠️ الاسم لا يمكن أن يكون فارغًا. الرجاء المحاولة مرة أخرى:");
    return;
  }
  
  userSessionManager.updateContext(chatId, { 
    name: name,
    waitingFor: null
  });
  
  var keyboard = {
    inline_keyboard: [
      [{ text: "👨‍💼 مشرف", callback_data: "add_user_permission_مشرف" }],
      [{ text: "👤 مستخدم عادي", callback_data: "add_user_permission_مستخدم عادي" }],
      [{ text: "❌ إلغاء", callback_data: "user_manage_main" }]
    ]
  };
  
  sendMessage(chatId, "🔑 الرجاء اختيار صلاحية المستخدم:", keyboard);
}

// Add a user to the main Users sheet
function addMainUser(chatId, userId, name, permission) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers);
  
  // Check if user already exists
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      // User exists, update the information
      sheet.getRange(i + 1, 2).setValue(permission); // Update permission
      sheet.getRange(i + 1, 3).setValue(name); // Update name
      
      sendMessage(chatId, "✅ تم تحديث معلومات المستخدم بنجاح.");
      userSessionManager.removeFromContext(chatId, ['userId', 'name']);
      return;
    }
  }
  
  // Add new user at the end of the sheet
  sheet.appendRow([userId, permission, name]);
  sendMessage(chatId, "✅ تمت إضافة المستخدم بنجاح.");
  userSessionManager.removeFromContext(chatId, ['userId', 'name']);
}

// Add a user to the broadcast list (Users2)
function addBroadcastUser(chatId, userId, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users2");
  
  // Check if user already exists
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      // User exists, update the name
      sheet.getRange(i + 1, 2).setValue(name);
      
      sendMessage(chatId, "✅ تم تحديث اسم المستخدم في قائمة البث بنجاح.");
      userSessionManager.removeFromContext(chatId, ['userId', 'name']);
      return;
    }
  }
  
  // Add new user at the end of the sheet
  sheet.appendRow([userId, name]);
  sendMessage(chatId, "✅ تمت إضافة المستخدم إلى قائمة البث بنجاح.");
  userSessionManager.removeFromContext(chatId, ['userId', 'name']);
}

// Process adding a new broadcast user - step 1 (chatId)
function processAddBroadcastUserStep1(chatId, text) {
  // Validate that the input is a valid chatId (number)
  var userId = text.trim();
  if (isNaN(userId)) {
    sendMessage(chatId, "⚠️ معرف المستخدم يجب أن يكون رقمًا. الرجاء المحاولة مرة أخرى:");
    return;
  }
  
  userSessionManager.updateContext(chatId, { 
    waitingFor: 'add_broadcast_user_name',
    userId: userId
  });
  
  sendMessage(chatId, "👤 الرجاء إرسال اسم المستخدم:");
}

// Process adding a new broadcast user - step 2 (name)
function processAddBroadcastUserStep2(chatId, text) {
  var name = text.trim();
  if (!name) {
    sendMessage(chatId, "⚠️ الاسم لا يمكن أن يكون فارغًا. الرجاء المحاولة مرة أخرى:");
    return;
  }
  
  var userId = userSessionManager.getSession(chatId).context.userId;
  addBroadcastUser(chatId, userId, name);
}

// Start editing a main user
function startEditMainUser(chatId, userId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers1);
  var data = sheet.getDataRange().getValues();
  var userRow = -1;
  var userData = null;
  
  // Find the user
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      userRow = i + 1;
      userData = data[i];
      break;
    }
  }
  
  if (userRow === -1) {
    sendMessage(chatId, "⚠️ لم يتم العثور على المستخدم.");
    return;
  }
  
  var keyboard = {
    inline_keyboard: [
      [{ text: "✏️ تعديل الاسم", callback_data: "edit_main_user_name_" + userId }],
      [{ text: "🔑 تغيير الصلاحية", callback_data: "edit_main_user_permission_" + userId }],
      [{ text: "❌ حذف المستخدم", callback_data: "delete_main_user_" + userId }],
      [{ text: "🔙 عودة", callback_data: "list_main_users" }]
    ]
  };
  
  var message = "👤 <b>تعديل المستخدم:</b>\n\n";
  message += "المعرف: " + userId + "\n";
  message += "الاسم: " + (userData[2] || "غير محدد") + "\n";
  message += "الصلاحية: " + (userData[1] || "مستخدم عادي") + "\n";
  
  sendMessage(chatId, message, keyboard);
}

// Start editing a main user's name
function startEditMainUserName(chatId, userId) {
  userSessionManager.updateContext(chatId, { 
    waitingFor: 'edit_main_user_name',
    editUserId: userId
  });
  
  sendMessage(chatId, "✏️ الرجاء إرسال الاسم الجديد للمستخدم:");
}

// Process editing a main user's name
function processEditMainUserName(chatId, text) {
  var userId = userSessionManager.getSession(chatId).context.editUserId;
  var newName = text.trim();
  
  if (!newName) {
    sendMessage(chatId, "⚠️ الاسم لا يمكن أن يكون فارغًا. الرجاء المحاولة مرة أخرى:");
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      sheet.getRange(i + 1, 3).setValue(newName);
      
      sendMessage(chatId, "✅ تم تحديث اسم المستخدم بنجاح.");
      userSessionManager.removeFromContext(chatId, ['waitingFor', 'editUserId']);
      return;
    }
  }
  
  sendMessage(chatId, "⚠️ لم يتم العثور على المستخدم.");
  userSessionManager.removeFromContext(chatId, ['waitingFor', 'editUserId']);
}

// Change a main user's permission
function changeMainUserPermission(chatId, userId, newPermission) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      sheet.getRange(i + 1, 2).setValue(newPermission);
      
      sendMessage(chatId, "✅ تم تحديث صلاحية المستخدم بنجاح.");
      return;
    }
  }
  
  sendMessage(chatId, "⚠️ لم يتم العثور على المستخدم.");
}

// Delete a main user
function deleteMainUser(chatId, userId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      sheet.deleteRow(i + 1);
      
      sendMessage(chatId, "✅ تم حذف المستخدم بنجاح.");
      return;
    }
  }
  
  sendMessage(chatId, "⚠️ لم يتم العثور على المستخدم.");
}

// Delete a broadcast user
function deleteBroadcastUser(chatId, userId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users2");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(userId)) {
      sheet.deleteRow(i + 1);
      
      sendMessage(chatId, "✅ تم حذف المستخدم من قائمة البث بنجاح.");
      return;
    }
  }
  
  sendMessage(chatId, "⚠️ لم يتم العثور على المستخدم.");
}




function toggleUserPermission(userId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetUsers1);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(userId)) {
      var currentPermission = data[i][3] || "مستخدم عادي";
      var newPermission = currentPermission === "مشرف" ? "مستخدم عادي" : "مشرف";
      sheet.getRange(i + 1, 4).setValue(newPermission); // العمود الرابع فيه الصلاحية
      return true;
    }
  }
  return false;
}
// ------------------------------------------------------------------ notification System

// ========== Updated Message Chunking Helper ==========
function chunkMessages(greeting, header, separator, rows) {
  const MAX_LENGTH = 4096;
  const preTag = "<pre>";
  const postTag = "</pre>";
  
  const baseMessage = `${greeting}${preTag}${header}\n${separator}\n`;
  let chunks = [];
  let currentChunk = [];
  let currentLength = baseMessage.length + postTag.length;

  rows.forEach(row => {
    const rowContent = `${row}`; // Already contains \n
    const potentialLength = currentLength + rowContent.length;

    if (potentialLength > MAX_LENGTH) {
      chunks.push(baseMessage + currentChunk.join("") + postTag);
      currentChunk = [rowContent];
      currentLength = baseMessage.length + postTag.length + rowContent.length;
    } else {
      currentChunk.push(rowContent);
      currentLength += rowContent.length;
    }
  });

  if (currentChunk.length > 0) {
    chunks.push(baseMessage + currentChunk.join("") + postTag);
  }

  return chunks;
}

function sendDailyTicketReport() {
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, "GMT+3", "yyyy-MM-dd");
  
  // Get the ActiveList sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ActiveList");
  if (!sheet) {
    Logger.log("ActiveList sheet not found");
    return;
  }
  
  // Get the data from the sheet
  var data = sheet.getDataRange().getValues();
  var headers = data[0]; // First row contains headers
  
  // Check if A2 is null
  var noTickets = !data[1] || !data[1][0];
  
  // Prepare the greeting
  var greeting = " 🌸 صباح الخير،\n" + "هي التذاكر موعدها قرّب\n\n";  
  if (noTickets) {
    var message = greeting + "لا توجد تذاكر تحتاج إلى مراجعة خلال الأيام السبعة القادمة.";
    sendToAllUsers(message);
  } else {
    var header = "اليوم هو: " + formattedDate + "\n\n";
    var separator = "-----";
    var rows = [];
    
    // Skip the header row and format each ticket
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Skip empty rows
      
      rows.push(separator + "\n");
      rows.push("تذكرة #" + (i) + ":\n");
      
      // Add each field with its header
      for (var j = 0; j < headers.length; j++) {
        if (data[i][j]) {
          rows.push(headers[j] + ": " + data[i][j] + "\n");
        }
      }
      rows.push("\n");
    }
    
    // Create chunked messages and send them
    var messageChunks = chunkMessages(greeting, header, separator, rows);
    sendChunksToAllUsers(messageChunks);
  }
  
  Logger.log("Daily report completed");
}

// Helper function to send to all users in Users2 sheet
function sendToAllUsers(message) {
  // Get the list of chat IDs from Users2 sheet
  var usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users2");
  if (!usersSheet) {
    Logger.log("Users2 sheet not found");
    return;
  }
  
  var chatIds = usersSheet.getRange("B2:B").getValues();
  
  // Send the message to each chat ID
  for (var i = 0; i < chatIds.length; i++) {
    var chatId = chatIds[i][0];
    if (chatId) {
      try {
        sendMessage(chatId, message);
        Logger.log("Message sent to " + chatId);
        Utilities.sleep(1000); // Add delay between messages
      } catch (error) {
        Logger.log("Error sending message to " + chatId + ": " + error);
      }
    }
  }
}

// Helper function to send chunks to all users
function sendChunksToAllUsers(chunks) {
  // Get the list of chat IDs from Users2 sheet
  var usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users2");
  if (!usersSheet) {
    Logger.log("Users2 sheet not found");
    return;
  }
  
  var chatIds = usersSheet.getRange("B2:B").getValues();
  
  // Send the message chunks to each chat ID
  for (var i = 0; i < chatIds.length; i++) {
    var chatId = chatIds[i][0];
    if (chatId) {
      try {
        for (var j = 0; j < chunks.length; j++) {
          sendMessage(chatId, chunks[j]);
          Utilities.sleep(1000); // Add delay between messages
        }
        Logger.log("All chunks sent to " + chatId);
      } catch (error) {
        Logger.log("Error sending message to " + chatId + ": " + error);
      }
    }
  }
}
