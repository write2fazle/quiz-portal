/**
 * দোকান হিসাব — Google Apps Script backend (Google Sheets সংস্করণ)
 * ------------------------------------------------------------
 * সবকিছু একটা Google Spreadsheet-এ আলাদা আলাদা শীটে (ট্যাব) রাখা হয়:
 *   Products, Purchases, Sales, Dues, Expenses, Contacts, OtherIncomes,
 *   Users, Settings, RecycleBin_Products, RecycleBin_Sales,
 *   RecycleBin_Purchases, RecycleBin_Dues
 *
 * ক্লায়েন্ট (HTML/JS) আগের মতোই পুরো ডাটা একসাথে getAllData() দিয়ে
 * পড়ে আর saveAllData() দিয়ে সেভ করে — তাই Index.html-এ কোনো পরিবর্তন
 * লাগে না, শুধু এই ফাইলটাই ডাটা কোথায় রাখবে সেটা ঠিক করে।
 */

const SPREADSHEET_NAME = 'দোকান হিসাব - ডাটাবেজ';
const SESSION_SECONDS = 6 * 60 * 60; // ৬ ঘণ্টা

const SHEETS = {
  PRODUCTS: 'Products',
  PURCHASES: 'Purchases',
  SALES: 'Sales',
  DUES: 'Dues',
  EXPENSES: 'Expenses',
  CONTACTS: 'Contacts',
  OTHER_INCOMES: 'OtherIncomes',
  USERS: 'Users',
  SETTINGS: 'Settings',
  RB_PRODUCTS: 'RecycleBin_Products',
  RB_SALES: 'RecycleBin_Sales',
  RB_PURCHASES: 'RecycleBin_Purchases',
  RB_DUES: 'RecycleBin_Dues'
};

const HEADERS = {
  PRODUCTS: ['id', 'name', 'category', 'unit', 'stock', 'lowStock', 'buyPrice', 'sellPrice'],
  PURCHASES: ['id', 'date', 'productId', 'qty', 'price', 'total', 'supplier', 'supplierPhone', 'contactId'],
  SALES: ['id', 'date', 'productId', 'qty', 'price', 'total', 'customer', 'customerPhone', 'contactId', 'payment', 'soldBy'],
  DUES: ['id', 'type', 'person', 'phone', 'contactId', 'date', 'total', 'paid', 'note'],
  EXPENSES: ['id', 'date', 'category', 'note', 'amount'],
  CONTACTS: ['id', 'type', 'name', 'phone', 'address', 'notes', 'avatarData'],
  OTHER_INCOMES: ['id', 'date', 'category', 'note', 'amount'],
  USERS: ['id', 'name', 'phone', 'address', 'email', 'avatarData', 'role', 'isOwner', 'permissions', 'salt', 'passwordHash'],
  SETTINGS: ['key', 'value']
};

const DEFAULT_ROLES = ['EMPLOYEE', 'CEO', 'Director', 'ED', 'MANAGER', 'OWNER'];

/* ================= DEBUG HELPERS ================= *
 * সমস্যা হলে Apps Script এডিটর থেকে সরাসরি এগুলো রান করে Logger-এ
 * ফলাফল দেখুন (View → Logs, অথবা Ctrl+Enter এর পর নিচে Execution log)।
 */

/** লগইন সরাসরি টেস্ট করুন — সঠিক এরর মেসেজ দেখতে এটা রান করুন। */
function debugLoginTest() {
  Logger.log(JSON.stringify(login('admin', 'admin123')));
}

/** স্প্রেডশিট পাওয়া যাচ্ছে কিনা ও Users শীটে কী আছে তা দেখুন। */
function debugInspectUsers() {
  const ss = requireSS_();
  Logger.log('Spreadsheet URL: ' + ss.getUrl());
  const users = readUsers_(ss);
  users.forEach(function (u) {
    Logger.log('phone=' + u.phone + ' | name=' + u.name + ' | isOwner=' + u.isOwner +
      ' | salt=' + u.salt + ' | passwordHash=' + u.passwordHash);
  });
  if (users.length === 0) Logger.log('Users শীটে কোনো ইউজার নেই!');
}

/** getAllData() ঠিকমতো একটা অবজেক্ট রিটার্ন করছে কিনা তা সরাসরি টেস্ট করুন। */
function debugGetAllDataTest() {
  const loginRes = login('admin', 'admin123');
  if (!loginRes.success) {
    Logger.log('লগইনই ব্যর্থ হয়েছে: ' + loginRes.message);
    return;
  }
  const data = getAllData(loginRes.token);
  Logger.log('getAllData() রিটার্ন করেছে: ' + (data ? 'একটা অবজেক্ট, ঠিক আছে ✅' : 'null/undefined — এটাই সমস্যা ❌'));
  if (data) {
    Logger.log('products=' + data.products.length + ' sales=' + data.sales.length +
      ' users=' + data.users.length + ' version=' + data.version);
  }
}

/**
 * কোনো Sheet/Session-এর উপর নির্ভর করে না — শুধু যাচাই করে যে
 * google.script.run ব্রিজ ও ডিপ্লয়মেন্ট আদৌ কাজ করছে কিনা।
 * getAllData() null আসলে ক্লায়েন্ট প্রথমে এটা কল করে দেখে।
 */
function ping() {
  return { ok: true, time: new Date().toString(), scriptId: ScriptApp.getScriptId() };
}

/**
 * ধাপে ধাপে প্রতিটা অংশ যাচাই করে একটা সরল (flat) অবজেক্ট রিটার্ন করে —
 * শুধু string/number/boolean, কোনো নেস্টেড অবজেক্ট/অ্যারে/Date নেই — যাতে
 * getAllData()-এ যদি সিরিয়ালাইজেশনে কোনো সমস্যা থাকে সেটা এড়িয়ে গিয়ে
 * ঠিক কোন ধাপে সমস্যা সেটা ধরা যায়।
 */
function debugDiagnose(token) {
  const out = { step: 'start', ok: true, detail: '' };
  try {
    out.step = 'requireSession_';
    requireSession_(token);

    out.step = 'requireSS_';
    const ss = requireSS_();
    out.spreadsheetUrl = ss.getUrl();

    out.step = 'readSettings_';
    const settings = readSettings_(ss);
    out.balance = settings.balance;
    out.version = settings.version;
    out.rolesCount = (settings.roles || []).length;

    const counts = {};
    ['PRODUCTS','PURCHASES','SALES','DUES','EXPENSES','CONTACTS','OTHER_INCOMES','USERS'].forEach(function (key) {
      out.step = 'readTable_:' + key;
      const sh = getSheet_(ss, SHEETS[key]);
      const rows = readTable_(sh, HEADERS[key]);
      counts[key] = rows.length;
    });
    out.counts = JSON.stringify(counts);

    out.step = 'readUsers_';
    const users = readUsers_(ss);
    out.usersCount = users.length;
    out.firstUserPhone = users.length ? users[0].phone : '(none)';

    out.step = 'sanitizeUser_';
    const sanitized = users.map(sanitizeUser_);
    out.sanitizedCount = sanitized.length;

    out.step = 'done';
    out.detail = 'সবগুলো ধাপ ঠিকভাবে শেষ হয়েছে।';
    return out;
  } catch (err) {
    out.ok = false;
    out.detail = String(err && err.message ? err.message : err);
    Logger.log('debugDiagnose FAILED at step "' + out.step + '": ' + out.detail);
    return out;
  }
}

/* ================= WEB APP ENTRY POINT ================= */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('দোকান হিসাব')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= ONE-TIME SETUP ================= *
 * Apps Script এডিটরে ফাংশন ড্রপডাউন থেকে "initialSetup" বেছে Run চাপুন।
 * এটা একটা নতুন Google Spreadsheet বানাবে ("দোকান হিসাব - ডাটাবেজ"),
 * দরকারি সব শীট/ট্যাব ও হেডার সারি তৈরি করবে, আর একটা ডিফল্ট মালিক
 * লগইন বানাবে:
 *     ইউজার আইডি : admin
 *     পাসওয়ার্ড   : admin123
 * ডিপ্লয়ের পর লগইন করেই পাসওয়ার্ড বদলে ফেলুন।
 */
function initialSetup() {
  const existing = findSpreadsheet_();
  if (existing) {
    Logger.log('স্প্রেডশিট ইতিমধ্যে আছে: ' + existing.getUrl());
    return;
  }

  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(SHEETS).forEach(function (key) {
    const name = SHEETS[key];
    const sh = ss.insertSheet(name);
    const headerKey = key.indexOf('RB_') === 0 ? key.substring(3) : key;
    const headers = HEADERS[headerKey] || HEADERS[key];
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  });
  // Apps Script always creates a default "Sheet1" — remove it now that our
  // named sheets exist.
  const def = ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);

  // Seed the owner login.
  const salt = Utilities.getUuid();
  const ownerRow = {
    id: Utilities.getUuid(), name: 'মালিক', phone: 'admin', address: '', email: '',
    avatarData: '', role: 'OWNER', isOwner: true, permissions: JSON.stringify([]),
    salt: salt, passwordHash: hash_('admin123', salt)
  };
  writeTable_(ss.getSheetByName(SHEETS.USERS), HEADERS.USERS, [ownerRow]);

  // Seed settings.
  writeTable_(ss.getSheetByName(SHEETS.SETTINGS), HEADERS.SETTINGS, [
    { key: 'balance', value: 0 },
    { key: 'roles', value: JSON.stringify(DEFAULT_ROLES) },
    { key: 'dataVersion', value: String(Date.now()) }
  ]);

  Logger.log('স্প্রেডশিট তৈরি হয়েছে: ' + ss.getUrl());
  Logger.log('ডিফল্ট লগইন -> ইউজার আইডি: admin / পাসওয়ার্ড: admin123 (প্রথম লগইনের পরই পরিবর্তন করুন!)');
}

/* ================= SPREADSHEET HELPERS (private) ================= */

function findSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (err) { /* stale id, fall through */ }
  }
  const it = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (it.hasNext()) {
    const f = it.next();
    props.setProperty('SPREADSHEET_ID', f.getId());
    return SpreadsheetApp.openById(f.getId());
  }
  return null;
}

function requireSS_() {
  const ss = findSpreadsheet_();
  if (!ss) throw new Error('স্প্রেডশিট পাওয়া যায়নি। Apps Script এডিটর থেকে প্রথমে initialSetup() ফাংশনটি একবার রান করুন।');
  return ss;
}

function getSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" নামে কোনো শীট পাওয়া যায়নি। initialSetup() আবার রান করে দেখুন।');
  return sh;
}

/** Reads all data rows of a sheet into an array of plain objects. */
function readTable_(sh, headers) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .filter(function (row) { return row.some(function (c) { return c !== '' && c !== null; }); })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

/** Clears a sheet's data rows and rewrites them from an array of objects. */
function writeTable_(sh, headers, objects) {
  const clearRows = Math.max(sh.getMaxRows() - 1, objects.length);
  if (clearRows > 0) sh.getRange(2, 1, clearRows, headers.length).clearContent();
  if (!objects || objects.length === 0) return;
  const rows = objects.map(function (o) {
    return headers.map(function (h) {
      const v = o[h];
      return (v === undefined || v === null) ? '' : v;
    });
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function safeParseArray_(v) {
  try {
    const p = JSON.parse(v);
    return Array.isArray(p) ? p : [];
  } catch (e) {
    return [];
  }
}

/* ---- Users (needs JSON <-> array + boolean coercion for permissions/isOwner) ---- */

function readUsers_(ss) {
  const raw = readTable_(getSheet_(ss, SHEETS.USERS), HEADERS.USERS);
  return raw.map(function (u) {
    return {
      id: u.id, name: u.name, phone: u.phone, address: u.address, email: u.email,
      avatarData: u.avatarData, role: u.role,
      isOwner: (u.isOwner === true || u.isOwner === 'TRUE' || u.isOwner === 'true'),
      permissions: safeParseArray_(u.permissions),
      salt: u.salt, passwordHash: u.passwordHash
    };
  });
}

function writeUsers_(ss, users) {
  const rows = users.map(function (u) {
    return {
      id: u.id, name: u.name, phone: u.phone, address: u.address || '', email: u.email || '',
      avatarData: u.avatarData || '', role: u.role, isOwner: !!u.isOwner,
      permissions: JSON.stringify(u.permissions || []), salt: u.salt, passwordHash: u.passwordHash
    };
  });
  writeTable_(getSheet_(ss, SHEETS.USERS), HEADERS.USERS, rows);
}

function sanitizeUser_(u) {
  return {
    id: u.id, name: u.name, phone: u.phone, address: u.address || '', email: u.email || '',
    avatarData: u.avatarData || '', role: u.role, isOwner: !!u.isOwner, permissions: u.permissions || []
  };
}

/* ---- Settings (balance + roles list) ---- */

function readSettings_(ss) {
  const rows = readTable_(getSheet_(ss, SHEETS.SETTINGS), HEADERS.SETTINGS);
  const map = {};
  rows.forEach(function (r) { map[r.key] = r.value; });
  const roles = safeParseArray_(map.roles);
  return {
    balance: Number(map.balance || 0),
    roles: roles.length ? roles : DEFAULT_ROLES.slice(),
    version: map.dataVersion ? String(map.dataVersion) : '0'
  };
}

function writeSettings_(ss, balance, roles, version) {
  writeTable_(getSheet_(ss, SHEETS.SETTINGS), HEADERS.SETTINGS, [
    { key: 'balance', value: Number(balance || 0) },
    { key: 'roles', value: JSON.stringify(roles && roles.length ? roles : DEFAULT_ROLES) },
    { key: 'dataVersion', value: String(version || Date.now()) }
  ]);
}

/* ---- Recycle bin (4 sub-sheets, mirrors client's {products,sales,purchases,dues}) ---- */

function readRecycleBin_(ss) {
  return {
    products: readTable_(getSheet_(ss, SHEETS.RB_PRODUCTS), HEADERS.PRODUCTS),
    sales: readTable_(getSheet_(ss, SHEETS.RB_SALES), HEADERS.SALES),
    purchases: readTable_(getSheet_(ss, SHEETS.RB_PURCHASES), HEADERS.PURCHASES),
    dues: readTable_(getSheet_(ss, SHEETS.RB_DUES), HEADERS.DUES)
  };
}

function writeRecycleBin_(ss, rb) {
  rb = rb || {};
  writeTable_(getSheet_(ss, SHEETS.RB_PRODUCTS), HEADERS.PRODUCTS, rb.products || []);
  writeTable_(getSheet_(ss, SHEETS.RB_SALES), HEADERS.SALES, rb.sales || []);
  writeTable_(getSheet_(ss, SHEETS.RB_PURCHASES), HEADERS.PURCHASES, rb.purchases || []);
  writeTable_(getSheet_(ss, SHEETS.RB_DUES), HEADERS.DUES, rb.dues || []);
}

/* ================= AUTH HELPERS (private) ================= */

function hash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + '::' + String(salt));
  return bytes.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function requireSession_(token) {
  if (!token) throw new Error('লগইন প্রয়োজন।');
  const raw = CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('সেশনের মেয়াদ শেষ হয়ে গেছে, আবার লগইন করুন।');
  return JSON.parse(raw); // { userId }
}

/* ================= PUBLIC API ================= *
 * ক্লায়েন্ট google.script.run.<functionName>(...) দিয়ে এগুলোই কল করে।
 */

function login(userId, password) {
  userId = String(userId || '').trim();
  const ss = requireSS_();
  const users = readUsers_(ss);
  const user = users.find(function (u) { return u.phone === userId; });
  if (!user) return { success: false, message: 'ইউজার আইডি খুঁজে পাওয়া যায়নি।' };

  const computed = hash_(password || '', user.salt);
  if (computed !== user.passwordHash) {
    return { success: false, message: 'পাসওয়ার্ড সঠিক নয়।' };
  }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify({ userId: user.id }), SESSION_SECONDS);
  return { success: true, token: token, user: sanitizeUser_(user) };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return { success: true };
}

function getAllData(token) {
  try {
    requireSession_(token);
    const ss = requireSS_();
    const settings = readSettings_(ss);
    const result = {
      products: readTable_(getSheet_(ss, SHEETS.PRODUCTS), HEADERS.PRODUCTS),
      purchases: readTable_(getSheet_(ss, SHEETS.PURCHASES), HEADERS.PURCHASES),
      sales: readTable_(getSheet_(ss, SHEETS.SALES), HEADERS.SALES),
      dues: readTable_(getSheet_(ss, SHEETS.DUES), HEADERS.DUES),
      expenses: readTable_(getSheet_(ss, SHEETS.EXPENSES), HEADERS.EXPENSES),
      contacts: readTable_(getSheet_(ss, SHEETS.CONTACTS), HEADERS.CONTACTS),
      otherIncomes: readTable_(getSheet_(ss, SHEETS.OTHER_INCOMES), HEADERS.OTHER_INCOMES),
      recycleBin: readRecycleBin_(ss),
      users: readUsers_(ss).map(sanitizeUser_),
      balance: settings.balance,
      roles: settings.roles,
      version: settings.version
    };
    Logger.log('getAllData() OK — products=' + result.products.length + ' users=' + result.users.length);
    return result;
  } catch (err) {
    Logger.log('getAllData() FAILED: ' + err + ' | stack: ' + (err && err.stack ? err.stack : 'n/a'));
    throw new Error('getAllData ব্যর্থ হয়েছে: ' + (err && err.message ? err.message : err));
  }
}

/**
 * Persists the client's full working copy of DATA into the spreadsheet.
 * Each array is written to its matching sheet; Users are merged specially
 * so password hashes (which the client never has) are preserved unless a
 * `newPassword` field is explicitly supplied.
 *
 * IMPORTANT — optimistic concurrency guard: the client must pass back the
 * `version` it received from getAllData(). If someone else has saved in
 * the meantime, the version on the server will have moved on, and this
 * save is REJECTED (nothing is written) instead of silently overwriting
 * newer data with a stale in-memory snapshot — this is what previously
 * caused users/products/etc. to appear to "auto-delete".
 */
function saveAllData(token, payloadJson, clientVersion) {
  requireSession_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('সার্ভার এই মুহূর্তে ব্যস্ত (আরেকজন সেভ করছেন), একটু পর আবার চেষ্টা করুন।');
  }
  try {
    const payload = JSON.parse(payloadJson);
    const ss = requireSS_();
    const settings = readSettings_(ss);

    if (clientVersion !== undefined && clientVersion !== null && clientVersion !== '' &&
        String(clientVersion) !== String(settings.version)) {
      return {
        success: false,
        conflict: true,
        message: 'অন্য কোনো ডিভাইস/ট্যাব থেকে এর মধ্যে ডাটা পরিবর্তন হয়ে গেছে। আপনার এই সেভটি বাতিল করা হলো যাতে সেই ডাটা মুছে না যায় — পাতাটি রিলোড করে আবার চেষ্টা করুন।',
        version: settings.version
      };
    }

    writeTable_(getSheet_(ss, SHEETS.PRODUCTS), HEADERS.PRODUCTS, payload.products || []);
    writeTable_(getSheet_(ss, SHEETS.PURCHASES), HEADERS.PURCHASES, payload.purchases || []);
    writeTable_(getSheet_(ss, SHEETS.SALES), HEADERS.SALES, payload.sales || []);
    writeTable_(getSheet_(ss, SHEETS.DUES), HEADERS.DUES, payload.dues || []);
    writeTable_(getSheet_(ss, SHEETS.EXPENSES), HEADERS.EXPENSES, payload.expenses || []);
    writeTable_(getSheet_(ss, SHEETS.CONTACTS), HEADERS.CONTACTS, payload.contacts || []);
    writeTable_(getSheet_(ss, SHEETS.OTHER_INCOMES), HEADERS.OTHER_INCOMES, payload.otherIncomes || []);
    writeRecycleBin_(ss, payload.recycleBin);

    const newVersion = String(Date.now());
    writeSettings_(ss, payload.balance, payload.roles, newVersion);

    const existingById = {};
    readUsers_(ss).forEach(function (u) { existingById[u.id] = u; });

    const mergedUsers = (payload.users || []).map(function (u) {
      const prev = existingById[u.id];
      let salt = prev ? prev.salt : Utilities.getUuid();
      let passwordHash = prev ? prev.passwordHash : null;

      if (u.newPassword) {
        salt = Utilities.getUuid();
        passwordHash = hash_(u.newPassword, salt);
      }
      if (!passwordHash) {
        // Safety net — a brand-new user must always get *some* password.
        salt = Utilities.getUuid();
        passwordHash = hash_(Utilities.getUuid(), salt);
      }

      return {
        id: u.id, name: u.name, phone: u.phone, address: u.address || '', email: u.email || '',
        avatarData: u.avatarData || '', role: u.role,
        isOwner: !!(prev && prev.isOwner), // owner status can never be granted from the client
        permissions: u.permissions || [], salt: salt, passwordHash: passwordHash
      };
    });
    writeUsers_(ss, mergedUsers);

    return { success: true, version: newVersion };
  } finally {
    lock.releaseLock();
  }
}

/** Lets a logged-in user change their own password. */
function changeMyPassword(token, currentPassword, newPassword) {
  const sess = requireSession_(token);
  const ss = requireSS_();
  const users = readUsers_(ss);
  const user = users.find(function (u) { return u.id === sess.userId; });
  if (!user) return { success: false, message: 'ইউজার পাওয়া যায়নি।' };

  if (hash_(currentPassword || '', user.salt) !== user.passwordHash) {
    return { success: false, message: 'বর্তমান পাসওয়ার্ড সঠিক নয়।' };
  }
  if (!newPassword || String(newPassword).length < 4) {
    return { success: false, message: 'নতুন পাসওয়ার্ড কমপক্ষে ৪ ক্যারেক্টার হতে হবে।' };
  }

  user.salt = Utilities.getUuid();
  user.passwordHash = hash_(newPassword, user.salt);
  writeUsers_(ss, users);
  return { success: true };
}
