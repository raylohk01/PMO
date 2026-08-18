/**
 * 📄 AuthLogic.gs
 * 使用者認證與權限檢查 (RBAC)
 */

// ==========================================
// 1. 權限矩陣定義 (Role-Access Matrix)
// ==========================================
// 根據 PRD，定義各部門 (Department) 的存取模組權限
const ROLE_ACCESS_MATRIX = {
  'Management': ['ALL'],
  'PM': ['Project List', 'Project Status', 'Project Overview'],
  'Editorial': ['Advertorial', 'Project Overview'],
  'Creative': ['Creative', 'Advertorial', 'Project Overview'],
  'Art': ['Creative', 'Advertorial', 'Event', 'Project Overview'],
  'Event': ['Event', 'Project Overview'],
  'Sales': ['Sales 模組', 'Project Overview'],
  'Marketing': ['Marketing 模組', 'Project Overview'],
  'Education': ['Education 模組', 'Project Overview']
};

// ==========================================
// 2. 核心讀取邏輯
// ==========================================
/**
 * 從 Users Sheet 讀取指定使用者的角色資料
 * @param {string} email - 使用者 Google 帳號
 * @returns {Object|null} - 回傳使用者物件，若找不到則回傳 null
 */
function getUserRole(email) {
  const users = getSheetData(SHEET_NAMES.USERS); // 呼叫 SheetHelper.gs 的函數
  
  // 尋找符合 email 且 isActive 不為 false 的使用者
  const user = users.find(u => u.email === email && String(u.isActive).toUpperCase() !== 'FALSE');
  
  if (user) {
    return {
      email: user.email,
      name: user.name,
      department: user.department,
      role: user.role
    };
  }
  return null;
}

/**
 * 回傳當前登入使用者的 email 與 role
 * @returns {Object} 
 */
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  
  if (!email) {
    throw new Error("系統無法獲取您的 Google 帳號，請確認執行權限。");
  }
  
  const user = getUserRole(email);
  if (!user) {
    throw new Error(`存取被拒絕：帳號 (${email}) 尚未在系統註冊或已被停權。`);
  }
  
  return user;
}

// ==========================================
// 3. 權限判定邏輯
// ==========================================
/**
 * 檢查當前使用者有沒有某個模組的存取權限
 * @param {string} moduleName - 模組名稱 (如 'Project Overview')
 * @returns {boolean}
 */
function hasAccess(moduleName) {
  const user = getCurrentUser();
  const allowedModules = ROLE_ACCESS_MATRIX[user.department] || [];
  
  if (allowedModules.includes('ALL')) return true; // Management 擁有全部權限
  return allowedModules.includes(moduleName);
}

/**
 * 檢查是否為 Management (最高管理層)
 * @returns {boolean}
 */
function isManagement() {
  return getCurrentUser().department === 'Management';
}

/**
 * 檢查是否為 PM 或 Head of PM
 * @returns {boolean}
 */
function isPM() {
  return getCurrentUser().department === 'PM';
}

/**
 * 檢查當前使用者是否可以看見某個專案
 * 規則：Management 和 PM 可看全部；其他團隊只能看自己被 assign 或相關的專案
 * @param {string} jobNumber - 專案編號
 * @returns {boolean}
 */
function canViewProject(jobNumber) {
  const user = getCurrentUser();
  
  // 1. 全局權限：Management 或是 PM 可以看所有專案
  if (isManagement() || isPM()) return true;
  
  // 2. 專案層級：檢查是否為該專案的相關負責人 (Sales, Editor, Art 等)
  const projects = getSheetData(SHEET_NAMES.PROJECTS);
  const project = projects.find(p => p.jobNumber === jobNumber);
  
  if (project) {
    const { name } = user; // 用中文名比對
    if (project.salesPerson === name || project.editorName === name || 
        project.copyName === name || project.artName === name) {
      return true;
    }
  }
  
  // 3. 任務層級：檢查 Tasks 表中是否有被 Assign 給這位 user
  const tasks = getSheetData(SHEET_NAMES.TASKS);
  const isAssignedToTask = tasks.some(t => t.jobNumber === jobNumber && t.assignedTo === user.email);
  
  return isAssignedToTask;
}

// ==========================================
// 4. 測試程式碼 (自我檢測用)
// ==========================================
function testAuthLogic() {
  const myEmail = Session.getActiveUser().getEmail();
  Logger.log('當前 Google 登入帳號: ' + myEmail);
  
  // 為了確保測試能成功，我們先把您的帳號強制作為 Management 寫入 Users 表
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.USERS);
  if (sheet.getLastRow() <= 1) {
    sheet.appendRow([myEmail, '系統管理員(測試)', 'Management', 'Management', 'TRUE']);
    Logger.log('已自動將您的帳號加入 Users 表，部門設為 Management！');
  }
  
  try {
    const user = getCurrentUser();
    Logger.log('👤 使用者資料讀取成功: ' + JSON.stringify(user));
    Logger.log('🛡️ 檢查是否為 Management: ' + isManagement());
    Logger.log('🛡️ 檢查是否為 PM: ' + isPM());
    Logger.log('🚪 檢查 Project Overview 權限: ' + hasAccess('Project Overview'));
    Logger.log('🚪 檢查 Event 模組權限: ' + hasAccess('Event 模組'));
  } catch (error) {
    Logger.log('❌ 測試發生錯誤: ' + error.message);
  }
}