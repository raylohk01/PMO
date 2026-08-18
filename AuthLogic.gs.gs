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
 * 從 Users Sheet 讀取指定使用者的角色資料 (嚴格防禦版)
 */
function getUserRole(email) {
  if (!email) return null;
  
  // 取得 Users 表資料
  const users = getSheetData(SHEET_NAMES.USERS) || [];
  const targetEmail = String(email).trim().toLowerCase();
  
  // 尋找符合 Email 的使用者
  const user = users.find(u => String(u.email || '').trim().toLowerCase() === targetEmail);
  
  if (user) {
    // 🛡️ 雙重檢查 status 與 isActive 欄位 (相容 Active/Inactive/FALSE/Disabled)
    const rawStatus = String(user.status || '').trim().toUpperCase();
    const rawActive = String(user.isActive || '').trim().toUpperCase();
    
    // 只要 status 或 isActive 包含 FALSE, INACTIVE, DISABLED 就代表已停權
    const isDisabled = (rawStatus === 'FALSE' || rawStatus === 'INACTIVE' || rawStatus === 'DISABLED') ||
                       (rawActive === 'FALSE' || rawActive === 'INACTIVE' || rawActive === 'DISABLED');
                       
    if (isDisabled) {
      Logger.log('⚠️ 帳號已停權: ' + email);
      return null; // 視同找不到該帳號，拒絕存取
    }

    return {
      email: user.email,
      name: user.name || user.email.split('@')[0],
      department: user.department || 'General',
      role: user.role || 'Member'
    };
  }
  
  return null;
}

/**
 * 獲取當前使用者 (嚴格驗證版)
 */
function getCurrentUser() {
  let email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch (e) {
    throw new Error('無法獲取您的 Google 登入資訊，請重新整理頁面。');
  }

  if (!email || email.trim() === '') {
    throw new Error('存取被拒絕：無法讀取您的 Google 帳號 Email。');
  }

  const user = getUserRole(email.trim());
  if (!user) {
    throw new Error(`⛔ 存取被拒絕：帳號 (${email}) 未在 Users 資料表中註冊，或該帳號已被停權 (Status: Disabled)。`);
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