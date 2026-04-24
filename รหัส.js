/**
 * SRI BROWN - B2B Subscription Backend (Google Apps Script)
 * ระบบจัดการโควตา (Allocation) และการเบิกกาแฟ (Withdrawal)
 * GitHub Repository: https://github.com/Issarapong-w/Sribrownsubscription
 */

const ADMIN_EMAIL = "Sribrowncafe@gmail.com";
const SHEET_ID = "1GAlB4_qMBKzrk3xdTt98rQihIBeduF0PM7_jX4jYpMw"; // Connected Sheet

const SHEET_CONFIG = {
  Users: ['email', 'password', 'company', 'phone', 'address', 'role', 'total_allocation', 'last_login'],
  Subscriptions: ['id', 'email', 'company', 'planName', 'price', 'billingCycle', 'status', 'date', 'amountKg'],
  Withdrawals: ['id', 'email', 'amount', 'status', 'timestamp', 'trackingNo'],
  Packages: ['id', 'name', 'price', 'kg', 'features', 'billingCycle', 'gradient', 'textColor']
};

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SHEET_ID);
  } catch (e) {
    console.error("Critical Error: Cannot open spreadsheet. Check SHEET_ID.");
    throw new Error("ระบบฐานข้อมูลขัดข้อง กรุณาตรวจสอบการตั้งค่า SHEET_ID");
  }
}

function findSheetByName(sheetName) {
  const ss = getSpreadsheet();
  const target = sheetName.toLowerCase();
  const sheets = ss.getSheets();
  return sheets.find(s => String(s.getName()).trim().toLowerCase() === target) || null;
}

function ensureSheet(sheetName) {
  const ss = getSpreadsheet();
  const headers = SHEET_CONFIG[sheetName];
  if (!headers) throw new Error('Unknown sheet: ' + sheetName);

  let sheet = findSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    return sheet;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0].map(h => String(h).trim());
  const needsReset = currentHeaders.length !== headers.length || currentHeaders.some((cell, idx) => cell !== headers[idx]);
  if (needsReset) {
    try {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } catch (e) {
      Logger.log("Failed to reset headers: " + e.message);
    }
  }
  return sheet;
}

function initializeSheets() {
  Object.keys(SHEET_CONFIG).forEach(ensureSheet);
  return { success: true, sheetNames: Object.keys(SHEET_CONFIG) };
}

function doGet() {
  ensureSheet('Users');
  ensureSheet('Subscriptions');
  ensureSheet('Withdrawals');
  ensureSheet('Packages');

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setTitle('SRI BROWN - Admin & Partner System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheetData(sheetName) {
  const sheet = ensureSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("No data in sheet: " + sheetName);
    return [];
  }
  
  const headers = data.shift().map(h => h.toString().trim().toLowerCase());
  const results = data.map((row, index) => {
    const obj = { rowid: index + 2 };
    headers.forEach((h, i) => {
      let value = row[i] !== undefined ? row[i] : "";
      // ตรวจสอบว่าเป็น Date หรือไม่ ถ้าใช่ให้แปลงเป็น String ป้องกัน Error ตอนส่งไป Frontend
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      }
      obj[h] = (value !== null && value !== undefined) ? value : "";
    });
    return obj;
  });
  Logger.log("Sheet " + sheetName + " data: " + results.length + " rows");
  return results;
}

// --- API Functions สำหรับลูกค้า ---

function checkLogin(email, password) {
  // Check sheet users
  const users = getSheetData('Users') || [];
  const cleanEmail = String(email || "").trim().toLowerCase();
  const cleanPassword = String(password || "").trim();

  const user = users.find(u => 
    String(u.email).toLowerCase() === cleanEmail && 
    String(u.password) === cleanPassword
  );

  if (user) {
    const { password, ...safeUser } = user;
    return safeUser;
  }
  return null;
}

function getUserProfile(email) {
  const users = getSheetData('Users');
  const user = users.find(u => u.email.toString().toLowerCase() === email.toString().trim().toLowerCase());
  if (user) {
    const { password, ...safeUser } = user;
    return safeUser;
  }
  return null;
}

/**
 * อัปเดตข้อมูลผู้ใช้ในคอลัมน์ที่กำหนด
 */
function _updateUserField(email, fieldName, value) {
  const sheet = ensureSheet('Users');
  const data = sheet.getDataRange().getValues();
  const colIdx = SHEET_CONFIG.Users.indexOf(fieldName) + 1;
  if (colIdx === 0) return false;

  const rowIndex = data.findIndex(row => String(row[0]).toLowerCase() === String(email).toLowerCase());
  if (rowIndex !== -1) {
    try {
      sheet.getRange(rowIndex + 1, colIdx).setValue(value);
    } catch (e) {
      Logger.log("Update failed: " + e.message);
      return false;
    }
    return true;
  }
  return false;
}

function registerUser(data) {
  const sheet = ensureSheet('Users');
  const role = data.email.toLowerCase() === ADMIN_EMAIL.toLowerCase() ? 'admin' : 'user';
  
  // สร้างแถวข้อมูลตามลำดับใน SHEET_CONFIG เพื่อความยืดหยุ่น
  const newUserRow = SHEET_CONFIG.Users.map(header => {
    if (header === 'role') return role;
    if (header === 'total_allocation') return 10; // Default 10kg
    if (header === 'last_login') return new Date();
    return data[header] || "";
  });

  sheet.appendRow(newUserRow); 
  return { success: true };
}

function getSubscriptions(email) {
  try {
    const all = getSheetData('Subscriptions');
    const targetEmail = String(email || "").trim().toLowerCase();
    if (!targetEmail) return [];

    const filtered = all.filter(s => {
      if (!s || !s.email) return false;
      return String(s.email).trim().toLowerCase() === targetEmail;
    });
    return filtered;
  } catch (e) {
    console.error("Error in getSubscriptions: " + e.message);
    return [];
  }
}

/**
 * ดึงข้อมูลการเบิกกาแฟตาม Email
 */
function getWithdrawals(email) {
  try {
    // 1. ดึงข้อมูลทั้งหมดจาก Sheet
    const all = getSheetData('Withdrawals'); 
    if (!all || !Array.isArray(all)) return [];

    // 2. ทำความสะอาด Email ที่รับมา
    const targetEmail = String(email || "").trim().toLowerCase();
    if (!targetEmail) return [];

    // 3. กรองข้อมูล (ใช้ Exact Match เพื่อความปลอดภัย)
    const filtered = all.filter(w => {
      if (!w || !w.email) return false;
      const withdrawalEmail = String(w.email).trim().toLowerCase();
      return withdrawalEmail === targetEmail;
    });

    // 4. ส่งข้อมูลกลับ (แนะนำให้เรียงลำดับจากใหม่ไปเก่าที่นี่เลยก็ได้ หรือไปทำที่ Frontend)
    return filtered;
  } catch (e) {
    console.error("Error in getWithdrawals: " + e.message);
    return [];
  }
}

function requestWithdrawal(email, amount) {
  const lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds for other processes to finish
    lock.waitLock(30000);

    const withdrawalSheet = ensureSheet('Withdrawals');
    const userSheet = ensureSheet('Users');
    const userData = getSheetData('Users');
    
    const user = userData.find(u => String(u.email).toLowerCase() === String(email).toLowerCase());
    
    if (!user) throw new Error("User not found");
    
    const currentAllocation = parseFloat(user.total_allocation || 0);
    const withdrawAmount = parseFloat(amount);

    if (isNaN(withdrawAmount) || withdrawAmount <= 0) {
      return { success: false, message: "Invalid withdrawal amount" };
    }

    if (currentAllocation < withdrawAmount) {
      return { success: false, message: "Insufficient coffee allocation balance" };
    }

    const id = "WD-" + new Date().getTime();
    const newAllocation = currentAllocation - withdrawAmount;
    
    // ใช้ Helper เพื่อความปลอดภัย
    _updateUserField(email, 'total_allocation', newAllocation);
    
    // Log the withdrawal
    withdrawalSheet.appendRow([id, email, withdrawAmount, 'pending', new Date()]);
    
    return { success: true, newBalance: newAllocation };
  
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// --- API Functions สำหรับ Admin ---

function getAllPartners() {
  return getSheetData('Users').filter(u => u.role !== 'admin');
}

function getAllWithdrawals() {
  const withdrawals = getSheetData('Withdrawals');
  const users = getSheetData('Users');
  
  return withdrawals.map(w => {
    // ค้นหาข้อมูลผู้ใช้ที่ตรงกับ email ในรายการเบิก
    const user = users.find(u => String(u.email).toLowerCase() === String(w.email).toLowerCase());
    return {
      ...w,
      company: user ? user.company : "ไม่พบชื่อร้าน",
      address: user ? user.address : "ไม่พบข้อมูลที่อยู่",
      phone: user ? user.phone : "ไม่พบเบอร์โทร"
    };
  });
}

function updateAllocation(email, newTotal) {
  const updated = _updateUserField(email, 'total_allocation', newTotal);
  if (updated) return { success: true };
  return { success: false };
}

function getPackages() {
  const sheet = ensureSheet('Packages');
  if (sheet.getLastRow() <= 1) {
    const initialData = [
      ['tiny', 'Tiny-pack', 1890, 3, 'เมล็ดกาแฟ 3 KG, Drip bag tester 3 bag', 'monthly', 'from-[#FAF9F6] to-[#F5F2F0]', 'text-[#3C2A21]'],
      ['medium', 'Medium-pack', 3100, 5, 'เมล็ดกาแฟ 5 KG, Drip bag tester 3 bag', 'monthly', 'from-[#F5F2F0] to-[#E8E2DE]', 'text-[#3C2A21]'],
      ['pro', 'Pro PACK', 6000, 10, 'เมล็ดกาแฟ 10 KG, Drip bag tester 5 bag', 'monthly', 'from-[#3C2A21] to-[#1A120B]', 'text-white'],
      ['starter', 'Starter', 21600, 36, 'เมล็ดกาแฟ 3 KG * 12 เดือน', 'annually', 'from-[#FAF9F6] to-[#F5F2F0]', 'text-[#3C2A21]'],
      ['growth', 'Growth', 35400, 60, 'เมล็ดกาแฟ 5 KG * 12 เดือน', 'annually', 'from-[#F5F2F0] to-[#E8E2DE]', 'text-[#3C2A21]'],
      ['enterprise', 'Organization', 68400, 120, 'เมล็ดกาแฟ 10 KG * 12 เดือน', 'annually', 'from-[#3C2A21] to-[#1A120B]', 'text-white']
    ];
    initialData.forEach(row => sheet.appendRow(row));
  }
  return getSheetData('Packages');
}

function addPackage(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureSheet('Packages');
    const id = "PKG-" + new Date().getTime();
    // เรียงลำดับตาม SHEET_CONFIG.Packages: ['id', 'name', 'price', 'kg', 'features', 'billingCycle', 'gradient', 'textColor']
    const row = [
      id, 
      data.name || "New Package", 
      data.price || 0, 
      data.kg || 0, 
      data.features || "Feature 1, Feature 2", 
      data.billingcycle || "monthly", 
      "from-[#FAF9F6] to-[#F5F2F0]", 
      "text-[#3C2A21]"
    ];
    sheet.appendRow(row);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updatePackage(rowId, field, value) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = findSheetByName('Packages');
    if (!sheet) return { success: false };
    const headers = SHEET_CONFIG.Packages;
    // ค้นหา Column Index แบบไม่สน Case เพื่อความปลอดภัย
    const colIdx = headers.findIndex(h => h.toLowerCase() === field.toLowerCase()) + 1;
    if (colIdx > 0) {
      sheet.getRange(rowId, colIdx).setValue(value);
      return { success: true };
    }
    return { success: false };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ยืนยันการชำระเงิน: เปลี่ยนสถานะ และบวกจำนวนกาแฟเข้าโควตาเดิม
 */
function confirmPayment(rowId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const subSheet = ensureSheet('Subscriptions');
    const userSheet = ensureSheet('Users');
    
    const subs = getSheetData('Subscriptions');
    const order = subs.find(s => s.rowid === rowId);
    
    if (!order || order.status === 'active') {
      return { success: false, message: "รายการนี้ถูกยืนยันไปแล้วหรือหาไม่พบ" };
    }

    // 1. อัปเดตสถานะ Order เป็น active
    const statusColIdx = SHEET_CONFIG.Subscriptions.indexOf('status') + 1;
    subSheet.getRange(rowId, statusColIdx).setValue('active');

    // 2. หา User เพื่อบวกยอดโควตา
    const userData = getSheetData('Users');
    const user = userData.find(u => u.email.toLowerCase() === order.email.toLowerCase());
    
    if (user) {
      const currentTotal = parseFloat(user.total_allocation || 0);
      const addAmount = parseFloat(order.amountkg || 0); // จากคอลัมน์ใหม่
      const newTotal = currentTotal + addAmount;
      
      const allocationColIdx = SHEET_CONFIG.Users.indexOf('total_allocation') + 1;
      userSheet.getRange(user.rowid, allocationColIdx).setValue(newTotal);
      
      return { success: true, message: `ยืนยันเรียบร้อย เติมโควตาใหม่เป็น ${newTotal} KG` };
    }
    
    return { success: false, message: "ไม่พบข้อมูลผู้ใช้งาน" };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ยืนยันการจัดส่งกาแฟ
 */
function confirmWithdrawal(rowId, trackingNo) {
  const sheet = findSheetByName('Withdrawals');
  if (sheet) {
    const statusColIdx = SHEET_CONFIG.Withdrawals.indexOf('status') + 1;
    const trackingColIdx = SHEET_CONFIG.Withdrawals.indexOf('trackingNo') + 1;
    sheet.getRange(rowId, statusColIdx).setValue('completed');
    sheet.getRange(rowId, trackingColIdx).setValue(trackingNo || "");
  }
  return { success: true };
}

function deleteRecord(sheetName, rowId) {
  const sheet = findSheetByName(sheetName);
  if (sheet) {
    sheet.deleteRow(rowId);
  }
  return { success: true };
}

function saveNewSubscription(data) {
  const lock = LockService.getScriptLock();
  try {
    // ล็อกสคริปต์เพื่อป้องกันปัญหา Race Condition
    lock.waitLock(15000);

    const ss = getSpreadsheet();
    const sheet = findSheetByName('Subscriptions') || ss.insertSheet('Subscriptions');
    
    const email = String(data.email).toLowerCase().trim();
    const values = sheet.getDataRange().getValues();
    const emailColIdx = SHEET_CONFIG.Subscriptions.indexOf('email');
    const statusColIdx = SHEET_CONFIG.Subscriptions.indexOf('status');

    // 1. ตรวจสอบและอัปเดตแพ็กเกจเดิมให้เป็น complete
    for (let i = 1; i < values.length; i++) {
      const rowEmail = String(values[i][emailColIdx]).toLowerCase().trim();
      const rowStatus = String(values[i][statusColIdx]).toLowerCase();
      
      if (rowEmail === email && rowStatus !== 'complete') {
        sheet.getRange(i + 1, statusColIdx + 1).setValue('complete');
      }
    }

    // 2. บันทึกแพ็กเกจใหม่
    const id = "SUB-" + new Date().getTime();
    sheet.appendRow([id, data.email, data.company, data.planName, data.price, data.billingCycle, 'pending', new Date(), data.amountKg]);
    
    return { success: true };
  } catch (e) {
    console.error("Error in saveNewSubscription: " + e.message);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getAllOrders() {
  return getSheetData('Subscriptions');
}