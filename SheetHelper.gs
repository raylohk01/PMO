/**
 * 📄 SheetHelper.gs
 * Google Sheets 資料庫結構定義與底層讀寫邏輯
 */

// ==========================================
// 1. 常數定義 (Constants & Enums)
// ==========================================
const SHEET_NAMES = {
  PROJECTS: 'Projects',
  USERS: 'Users',
  TASKS: 'Tasks',
  INDUSTRIES: 'Industries',
  AUDIT_LOG: 'AuditLog',
  CONFIG: 'Config'
};

const COLUMNS = {
  PROJECTS: [
    'jobNumber', 'submissionDate', 'launchDate', 'jobNature', 'jobStatus', 'soStatus',
    'clientName', 'productName', 'textJobType', 'itemType', 'salesPerson', 'pmName',
    'editorName', 'copyName', 'artName',
    // 數字明細欄位
    'boostingBudget', 'videoBuyout', 'advertorialArticle', 'advertorialFeed', 'editorialArticle', 
    'editorialFeed', 'outing', 'video', 'live', 'bannerAd', 'appPush', 'whatsapp', 
    'linkedin', 'xiaohongshu', 'edm', 'booklet', 'event', 'pitchingDeck', 'pitchingQuotation', 
    'issuePage', 'minisite',
    // 狀態與系統欄位
    'industry', 'status', 'firstDraftDate', 'completedDate', 'versionNo', 
    'googleDriveFolderUrl', 'createdAt', 'updatedAt'
  ],
  USERS: ['email', 'name', 'department', 'role', 'isActive'],
  TASKS: [
    'taskId', 'jobNumber', 'taskType', 'assignedTo', 'status', 'deadline', 
    'outline', 'firstDraftUrl', 'revisedUrl', 'designDraftUrl', 'revisedDesignUrl', 
    'acceptDate', 'completedDate', 'versionNo'
  ],
  INDUSTRIES: ['行業名稱'],
  AUDIT_LOG: ['時間', '使用者', '動作', '專案編號', '詳細內容'],
  CONFIG: ['設定名稱', '設定值']
};

const DEFAULT_INDUSTRIES = [
  "Airline", "Art & Culture", "Auction", "Automotive & Petroleum", "AV Products", 
  "Banking & Finance", "Beauty/Cosmetics/Skincare & Fitness", "Business Service", 
  "Camera/Photography & Optical Instruments", "Chamber", "Corporate", "Education", 
  "Entertainment", "Exhibition", "Fashion", "FMCG", "Food & Beverage", "Gadget/Game & Platform", 
  "Government", "Home & Living", "Home Appliance", "Infant Milk Formula", "Insurance", "Media", 
  "Miscellaneous", "NGO", "Optical & Eyewear", "Overseas Property", "Pharmaceuticals & Healthcare", 
  "Property & Shopping Mall", "Retail", "Sports", "Telecom/Computer & Internet Service", 
  "Travel", "Watch & Jewellery"
];

// ==========================================
// 2. 初始化函數 (Initialize Sheets)
// ==========================================
/**
 * 一鍵建立 6 張 Sheets 並設定標題列與預設值
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 建立並設定標題列
  Object.keys(SHEET_NAMES).forEach(key => {
    const sheetName = SHEET_NAMES[key];
    const headers = COLUMNS[key];
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // 設定第一行標題、凍結與加粗
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(1, 1, 1, headers.length).setBackground('#f3f3f3');
    }
  });
  
  // 寫入 Industries 預設值
  const industrySheet = ss.getSheetByName(SHEET_NAMES.INDUSTRIES);
  if (industrySheet.getLastRow() <= 1) {
    const industryData = DEFAULT_INDUSTRIES.map(ind => [ind]);
    industrySheet.getRange(2, 1, industryData.length, 1).setValues(industryData);
  }
  
  // 寫入 Config 預設值
  const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (configSheet.getLastRow() <= 1) {
    const configData = [
      ['SLA_DESIGN_DAYS', '1'],
      ['SLA_EDITORIAL_DAYS', '3'],
      ['MAX_VERSION', '5']
    ];
    configSheet.getRange(2, 1, configData.length, 2).setValues(configData);
  }
  
  Logger.log('✅ 所有資料庫 Sheets 初始化完成！');
}

// ==========================================
// 3. 基礎讀寫輔助函數 (Helper Functions)
// ==========================================

/**
 * 將整張 Sheet 的資料讀取為 Array of Objects
 * @param {string} sheetName - Sheet 名稱
 * @returns {Array<Object>} 
 */
function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // 只有標題或空白
  
  const headers = COLUMNS[Object.keys(SHEET_NAMES).find(k => SHEET_NAMES[k] === sheetName)] || data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

/**
 * 新增一筆資料 (根據傳入的 Object 自動對應欄位寫入)
 * @param {string} sheetName - Sheet 名稱
 * @param {Object} rowDataObj - 要寫入的資料物件
 */
function appendRow(sheetName, rowDataObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet ${sheetName} not found.`);
  
  const headers = COLUMNS[Object.keys(SHEET_NAMES).find(k => SHEET_NAMES[k] === sheetName)];
  
  const rowArray = headers.map(header => {
    // 針對數字欄位給予預設值 0
    if (sheetName === SHEET_NAMES.PROJECTS && header !== 'jobNumber' && typeof rowDataObj[header] === 'undefined') {
       const numberFields = ['boostingBudget', 'videoBuyout', 'advertorialArticle', 'advertorialFeed', 'editorialArticle', 'editorialFeed', 'outing', 'video', 'live', 'bannerAd', 'appPush', 'whatsapp', 'linkedin', 'xiaohongshu', 'edm', 'booklet', 'event', 'pitchingDeck', 'pitchingQuotation', 'issuePage', 'minisite', 'versionNo'];
       if (numberFields.includes(header)) return 0;
    }
    return rowDataObj[header] !== undefined ? rowDataObj[header] : '';
  });
  
  sheet.appendRow(rowArray);
}

/**
 * 更新單筆資料 (透過 PK 尋找並更新對應欄位)
 * @param {string} sheetName - Sheet 名稱
 * @param {string} pkColumnName - Primary Key 的欄位名稱 (如 'jobNumber', 'email')
 * @param {string} pkValue - Primary Key 的值
 * @param {Object} updateDataObj - 要更新的資料物件
 * @returns {boolean} 是否更新成功
 */
function updateRow(sheetName, pkColumnName, pkValue, updateDataObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet ${sheetName} not found.`);
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  
  const headers = COLUMNS[Object.keys(SHEET_NAMES).find(k => SHEET_NAMES[k] === sheetName)];
  const pkIndex = headers.indexOf(pkColumnName);
  
  if (pkIndex === -1) throw new Error(`Primary Key ${pkColumnName} not found in headers.`);
  
  // 尋找目標列 (Row Index 是 0-based，Spreadsheet row 是 1-based，且要跳過標題)
  let targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][pkIndex] === pkValue) {
      targetRowIndex = i + 1; // 轉為 Spreadsheet row number
      break;
    }
  }
  
  if (targetRowIndex === -1) return false; // 找不到對應的 PK
  
  // 準備更新資料
  Object.keys(updateDataObj).forEach(key => {
    const colIndex = headers.indexOf(key);
    if (colIndex !== -1) {
      sheet.getRange(targetRowIndex, colIndex + 1).setValue(updateDataObj[key]);
    }
  });
  
  return true;
}

/**
 * 寫入 AuditLog 系統操作紀錄
 * @param {string} action - 動作名稱 (如 CREATE_PROJECT)
 * @param {string} jobNumber - 關聯的專案編號
 * @param {string} details - 詳細內容
 */
function logAction(action, jobNumber, details) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.AUDIT_LOG);
    if (!sheet) return;
    const user = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    sheet.appendRow([timestamp, user, action, jobNumber, details]);
  } catch (e) {
    Logger.log('寫入 Log 失敗: ' + e.message);
  }
}