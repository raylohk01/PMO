/**
 * 📄 Code.gs
 * Web App 主入口與 API Controller (總控台)
 * 負責處理權限驗證、錯誤攔截與前後端資料交換
 */

// ==========================================
// 1. Web App 入口 (doGet) - 第一線大樓門禁
// ==========================================
function doGet(e) {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    return HtmlService.createHtmlOutput('<h1>存取被拒</h1><p>請先登入 Google 帳號。</p>');
  }

  // 🛡️ 大樓門禁：開啟網頁瞬間查驗 Users 表，未註冊或停權直接擋在門外
  try {
    const user = getUserRole(email);
    if (!user) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif; text-align:center; padding:50px;">' +
          '<h1 style="color:#dc3545; font-size:3em;">⛔ 存取被拒絕</h1>' +
          '<p style="font-size:1.2em; color:#555;">帳號 (' + email + ') 尚未在 Users 資料表中註冊，或該帳號已被系統停用 (Status: Inactive/FALSE)。</p>' +
          '<p style="color:#888;">請聯絡 PMO 管理員開通權限。</p>' +
        '</div>'
      );
    }
  } catch (err) {
    return HtmlService.createHtmlOutput('<h1>系統錯誤</h1><p>' + err.message + '</p>');
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('HK01 PMO 專案管理系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// 2. 權限驗證輔助函數 (verifyAuth) - 房間門鎖 (務必保留！)
// ==========================================
/**
 * 驗證使用者是否具備指定權限
 * @param {Array<string>} allowedRoles - 允許的角色陣列，如 ['Admin', 'PM']
 * @returns {Object} 使用者資訊
 */
function verifyAuth(allowedRoles = []) {
  // 🛡️ 呼叫 AuthLogic.gs 的 getCurrentUser，確保停權者在執行 API 時同樣被鎖定
  const user = getCurrentUser(); 
  
  if (allowedRoles.length > 0 && !allowedRoles.includes(user.role)) {
    throw new Error(`權限不足！此操作需要以下角色之一：${allowedRoles.join(', ')}`);
  }
  return user;
}

// ==========================================
// 3. API Endpoints
// ==========================================

function api_getCurrentUser() {
  try {
    const user = verifyAuth();
    return { success: true, data: user };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 支援帳號模擬的 API 接口
// ==========================================

function api_getDashboardData(simulateEmail) {
  try {
    const user = verifyAuth();
    // 💡 核心：如果有傳入 simulateEmail 就用它，否則用真實登入者的 email
    const targetEmail = simulateEmail || user.email; 
    
    // 依賴您原本寫好的 ViewLogic，直接把目標信箱丟進去！
    const data = getIndividualView(targetEmail);
    return { success: true, data: data.data };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/*
function api_getClientReviewList(simulateEmail) {
  try {
    const user = verifyAuth(['Sales', 'PM', 'Admin', 'Management']);
    const targetEmail = simulateEmail || user.email;
    
    // 這裡假設您有對應的 ViewLogic 函數，若沒有，可沿用之前寫好的抓取邏輯
    // 只要確保過濾條件是使用 targetEmail 即可
    const data = getClientReviewView(targetEmail); 
    return { success: true, data: data.data };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
*/

function api_getWorkloadStats(department, simulateEmail) {
  try {
    verifyAuth(['Admin', 'Management', 'Team Head']);
    
    const tasks = getSheetData(SHEET_NAMES.TASKS) || [];
    const users = getSheetData(SHEET_NAMES.USERS) || [];

    // 若有傳入 department 則優先使用，否則預設 Editorial
    const targetDept = department || 'Editorial';

    const inProgress = tasks.filter(t => t.status !== 'Completed' && t.status !== 'Cancelled');
    
    const workload = {};
    inProgress.forEach(t => {
      const pic = t.assignedTo || 'Unassigned';
      if (!workload[pic]) workload[pic] = { count: 0 };
      workload[pic].count++;
    });

    return { 
      success: true, 
      data: { 
        inProgress: inProgress.map(t => ({
          jobNumber: t.jobNumber || 'N/A',
          client: 'HK01 Project',
          PIC: t.assignedTo || 'Unassigned',
          status: t.status || 'In Progress',
          deadline: t.deadline ? String(t.deadline).split('T')[0] : 'N/A'
        })),
        workload: workload 
      } 
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getReports(simulateEmail) {
  try {
    verifyAuth(['Admin', 'Management']);
    
    const projects = getSheetData(SHEET_NAMES.PROJECTS) || [];
    const tasks = getSheetData(SHEET_NAMES.TASKS) || [];
    
    const highRiskProjects = [];
    const deptStats = { 'Editorial': { blocked: 0, overdue: 0 }, 'Creative': { blocked: 0, overdue: 0 }, 'PM': { blocked: 0, overdue: 0 } };

    projects.forEach(p => {
      if (p.status === 'Blocked') {
        highRiskProjects.push({
          jobNumber: p.jobNumber,
          client: p.clientName || p.client || 'N/A',
          reason: '專案狀態處於 Blocked (卡關)',
          PIC: p.pmName || 'PM'
        });
        if (deptStats['PM']) deptStats['PM'].blocked++;
      }
    });

    tasks.forEach(t => {
      if (t.status === 'Blocked') {
        highRiskProjects.push({
          jobNumber: t.jobNumber,
          client: 'Task Level',
          reason: `任務 [${t.taskType}] 處於 Blocked`,
          PIC: t.assignedTo
        });
        if (deptStats['Editorial']) deptStats['Editorial'].blocked++;
      }
    });

    return { success: true, data: { highRiskProjects: highRiskProjects, deptStats: deptStats } };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getProjectList(filter = {}) {
  try {
    verifyAuth(['Admin', 'Management', 'Team Head', 'PM', 'Sales']);
    let projects = getSheetData(SHEET_NAMES.PROJECTS);
    
    // 簡易過濾器實作
    if (filter.status) projects = projects.filter(p => p.status === filter.status);
    if (filter.pmName) projects = projects.filter(p => p.pmName === filter.pmName);
    
    return { success: true, data: projects };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getProjectDetail(jobNumber) {
  try {
    verifyAuth();
    const project = getProject(jobNumber);
    if (!project) throw new Error('找不到專案');
    return { success: true, data: project };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_updateProject(jobNumber, updates) {
  try {
    verifyAuth(['Admin', 'PM']);
    updateProjectStatus(jobNumber, updates.status); // 簡化版：目前僅示範更新狀態
    return { success: true, message: '專案更新成功' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_startProject(jobNumber) {
  try {
    verifyAuth(['Admin', 'PM']);
    startProject(jobNumber); // 呼叫 ProjectLogic
    return { success: true, message: '專案已成功啟動' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_assignTask(data) {
  try {
    verifyAuth(['Admin', 'PM', 'Team Head']);
    // data: { jobNumber, taskType, assignedTo, deadline, requiresDrive }
    const result = assignTask(data.jobNumber, data.taskType, data.assignedTo, data.deadline, data.requiresDrive);
    return { success: true, data: result, message: '任務指派成功' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_acceptTask(jobNumber, taskId) {
  try {
    verifyAuth();
    // 將狀態從 Waiting for PIC 改為 In Progress
    updateTaskStatus(taskId, 'In Progress');
    return { success: true, message: '已接受任務並開始處理' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_submitFirstDraft(jobNumber, taskId, url) {
  try {
    verifyAuth();
    submitTaskDraft(taskId, url, true); // true 代表是初稿
    return { success: true, message: '初稿已提交' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_submitRevised(jobNumber, taskId, url) {
  try {
    verifyAuth();
    submitTaskDraft(taskId, url, false); // false 代表是修改稿
    return { success: true, message: '修改稿已提交' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_completeProject(jobNumber) {
  try {
    verifyAuth(['Admin', 'PM']);
    completeProject(jobNumber);
    return { success: true, message: '專案已結案' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/*
function api_getClientReviewList() {
  try {
    verifyAuth(['Admin', 'PM', 'Sales', 'Management']);
    const result = getClientReviewList(); // 呼叫 ClientReviewLogic
    if (!result.success) throw new Error(result.message);
    return { success: true, data: result.data };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
*/

function api_sendReminder(jobNumber, method) {
  try {
    verifyAuth(['Admin', 'PM', 'Sales']);
    const result = sendClientReminder(jobNumber, method);
    return { success: true, message: result.message };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getReports() {
  try {
    verifyAuth(['Admin', 'Management']);
    
    // 1. 抓取高風險專案 (Blocked 或 逾期)
    const projects = getSheetData(SHEET_NAMES.PROJECTS);
    const tasks = getSheetData(SHEET_NAMES.TASKS);
    
    const highRiskProjects = [];
    const deptStats = { 'Editorial': { blocked: 0, overdue: 0 }, 'Creative': { blocked: 0, overdue: 0 }, 'PM': { blocked: 0, overdue: 0 } };

    projects.forEach(p => {
      if (p.status === 'Blocked') {
        highRiskProjects.push({
          jobNumber: p.jobNumber,
          client: p.clientName || p.client || 'N/A',
          reason: '專案狀態處於 Blocked (卡關)',
          PIC: p.pmName || 'PM'
        });
        if (deptStats['PM']) deptStats['PM'].blocked++;
      }
    });

    tasks.forEach(t => {
      if (t.status === 'Blocked') {
        highRiskProjects.push({
          jobNumber: t.jobNumber,
          client: 'Task Level',
          reason: `任務 [${t.taskType}] 處於 Blocked`,
          PIC: t.assignedTo
        });
        if (deptStats['Editorial']) deptStats['Editorial'].blocked++;
      }
    });

    return { 
      success: true, 
      data: { 
        highRiskProjects: highRiskProjects, 
        deptStats: deptStats 
      } 
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getWorkloadStats(department) {
  try {
    verifyAuth(['Admin', 'Management', 'Team Head']);
    
    // 取得所有任務與使用者
    const tasks = getSheetData(SHEET_NAMES.TASKS) || [];
    const users = getSheetData(SHEET_NAMES.USERS) || [];

    // 預設抓取 Editorial，若有指定則用指定的
    const targetDept = department || 'Editorial';

    // 找出屬於該部門的所有 Users (比對 Email 或 Name)
    const deptUserEmails = users
      .filter(u => u.department === targetDept || targetDept === 'Editorial') // 測試時允許彈性比對
      .map(u => (u.email || '').toLowerCase());

    // 篩選出進行中或卡關的任務
    const inProgress = tasks.filter(t => t.status !== 'Completed' && t.status !== 'Cancelled');
    
    const workload = {};
    inProgress.forEach(t => {
      const pic = t.assignedTo || 'Unassigned';
      if (!workload[pic]) workload[pic] = { count: 0 };
      workload[pic].count++;
    });

    return { 
      success: true, 
      data: { 
        inProgress: inProgress.map(t => ({
          jobNumber: t.jobNumber || 'N/A',
          client: 'HK01 Project',
          PIC: t.assignedTo || 'Unassigned',
          status: t.status || 'In Progress',
          deadline: t.deadline ? String(t.deadline).split('T')[0] : 'N/A'
        })),
        workload: workload 
      } 
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}


// ==========================================
// [Code 18-A] HTML 模組化拆檔引擎
// ==========================================

/**
 * 允許在主 HTML 中引入其他的 HTML 檔案
 * 用法：在 Index.html 中寫 <?!= include('檔案名稱'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// [工作流詳細資料 API] 補齊前端呼叫需要的函數 (相容 logs 與 auditLog)
// ==========================================
function api_getProjectWorkflowDetails(jobNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) throw new Error('無專案數據');

    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      
      if (currentJob.toLowerCase() === String(jobNumber).toLowerCase().trim()) {
        const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '--';
        const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '--';
        const status = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : 'In Progress';

        let wfData = {};
        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.includes('deliverables') || cellStr.includes('auditLog')) {
            try { wfData = JSON.parse(cellStr); break; } catch(e){}
          }
        }

        // 💡 雙重支援：優先抓取 auditLog，次選 logs
        const finalLogs = wfData.auditLog || wfData.logs || [];

        return {
          success: true,
          project: {
            jobNumber: currentJob,
            client: clientName,
            pmName: pmName,
            status: status,
            deliverables: wfData.deliverables || (Array.isArray(wfData) ? wfData : []),
            logs: finalLogs,
            auditLog: finalLogs
          }
        };
      }
    }

    throw new Error('找不到編號為 ' + jobNumber + ' 的專案');

  } catch (e) {
    return { success: false, message: e.message };
  }
}



// ==========================================
// [強健版 API] 專案整體管理狀態更新 (含完整防禦機制)
// ==========================================
function api_manageProjectStatus(payloadInput, optionalAction) {
  try {
    // 🛡️ 防禦 1：檢查基本輸入參數是否存在
    if (!payloadInput && !optionalAction) {
      throw new Error('未傳入任何有效參數');
    }

    let targetJobNumber = '';
    let action = '';

    // 🛡️ 防禦 2：安全解析 JSON，防止前端傳入不符合格式的字串導致程序崩潰
    if (typeof payloadInput === 'string' && payloadInput.trim().startsWith('{')) {
      try {
        let parsed = JSON.parse(payloadInput);
        targetJobNumber = String(parsed.jobNumber || parsed.jobNo || parsed.id || '').trim();
        action = String(parsed.action || '').trim();
      } catch(e) {
        Logger.log('⚠️ JSON 解析警告: ' + e.message);
      }
    }

    if (!targetJobNumber) {
      targetJobNumber = String(payloadInput || '').trim();
      action = String(optionalAction || '').trim();
    }

    if (!targetJobNumber || targetJobNumber === '[object Object]') {
      throw new Error('無法辨識有效的專案編號 (Job Number)');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('無法連線至試算表資料庫');
    
    const sheet = ss.getSheetByName('Projects');
    if (!sheet) throw new Error('資料庫缺少 [Projects] 工作表');

    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) throw new Error('Projects 工作表中無任何資料');

    // 🛡️ 防禦 3：標題列欄位防護與安全清洗
    const headersLower = data[0].map(h => String(h || '').trim().toLowerCase());
    
    const idxJobNum = headersLower.findIndex(h => h.includes('jobnumber') || h === 'jobno' || h === 'id');
    const idxStatus = headersLower.findIndex(h => h === 'status' || h === 'project_status');
    const idxTextJob = headersLower.findIndex(h => h === 'textjobtype');

    if (idxJobNum === -1) throw new Error('工作表欄位缺失: 找不到 JobNumber 相關欄位');
    if (idxStatus === -1) throw new Error('工作表欄位缺失: 找不到 Status 欄位');
    if (idxTextJob === -1) throw new Error('工作表欄位缺失: 找不到 textJobType 欄位');

    // 🛡️ 防禦 4：安全獲取當前使用者名稱
    let currentUser = 'Unknown_User';
    try {
      const activeUser = Session.getActiveUser();
      if (activeUser && activeUser.getEmail()) {
        currentUser = activeUser.getEmail().split('@')[0];
      }
    } catch(e) {
      Logger.log('無法獲取 Session 用戶，使用預設身份: ' + e.message);
    }

    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Hong_Kong', 'yyyy-MM-dd HH:mm');
    let newStatus = '', actionType = '', logDetail = '';

    const actUpper = String(action).toUpperCase();
    if (actUpper === 'PAUSE') {
      newStatus = 'Paused';
      actionType = 'Pause Project';
      logDetail = '⏸️ 將整體專案設為【暫停】';
    } else if (actUpper === 'RESUME') {
      newStatus = 'In Progress';
      actionType = 'Resume Project';
      logDetail = '▶️ 將整體專案【恢復執行】';
    } else if (actUpper === 'DELETE') {
      newStatus = 'Recycle Bin';
      actionType = 'Delete Project';
      logDetail = '🗑️ 將整體專案移至【回收箱】';
    } else {
      throw new Error('不支援的操作類型：' + action);
    }

    const newLogEntry = {
      timestamp: nowStr,
      user: currentUser,
      action: actionType,
      details: logDetail
    };

    // 3. 搜尋並更新專案
    let isUpdated = false;
    for (let i = 1; i < data.length; i++) {
      const currentJob = String(data[i][idxJobNum] || '').trim();

      if (currentJob.toLowerCase() === targetJobNumber.toLowerCase()) {
        // 更新狀態
        sheet.getRange(i + 1, idxStatus + 1).setValue(newStatus);

        // 🛡️ 防禦 5：Log 安全讀取與防壞軌 JSON
        let cellValue = String(data[i][idxTextJob] || '').trim();
        let logArray = [];
        
        if (cellValue.startsWith('[')) {
          try {
            logArray = JSON.parse(cellValue);
          } catch(e) {
            Logger.log('⚠️ 歷史 Log JSON 解析失敗，重置為新陣列');
            logArray = [];
          }
        }
        
        if (!Array.isArray(logArray)) logArray = [];
        
        logArray.unshift(newLogEntry); // 最新紀錄至最前
        sheet.getRange(i + 1, idxTextJob + 1).setValue(JSON.stringify(logArray));
        
        isUpdated = true;
        break;
      }
    }

    if (!isUpdated) {
      throw new Error('資料庫中找不到編號為 [' + targetJobNumber + '] 的專案');
    }

    return {
      success: true,
      message: '專案狀態已成功更新為 [' + newStatus + ']！'
    };

  } catch (e) {
    // 🛡️ 防禦 6：統一捕捉結構化 Error 避免前端程序跳出未捕捉例外
    return { success: false, message: '❌ 操作被系統攔截: ' + e.message };
  }
}

function api_saveUser(input) {
  try {
    verifyAuth(['Admin', 'Management']);

    // 💡 捕捉傳進來的原始資料型態與內容
    const rawInputType = Array.isArray(input) ? 'Array (陣列)' : typeof input;
    const rawInputJson = JSON.stringify(input);

    let empId = '', email = '', name = '', department = '', role = '', status = 'Active';

    if (Array.isArray(input)) {
      empId = String(input[0] || '').trim();
      email = String(input[1] || '').trim();
      name = String(input[2] || '').trim();
      department = String(input[3] || '').trim();
      role = String(input[4] || '').trim();
      status = String(input[5] || 'Active').trim();
    } else if (input && typeof input === 'object') {
      empId = String(input.empId || input.empID || '').trim();
      email = String(input.email || '').trim();
      name = String(input.name || '').trim();
      department = String(input.department || input.dept || '').trim();
      role = String(input.role || '').trim();
      status = String(input.status || 'Active').trim();
    }

    const isActive = (status === 'Active');

    // 精準 A ~ G 欄位陣列
    const rowData = [
      email,        // A 欄
      name,         // B 欄
      department,   // C 欄
      role,         // D 欄
      isActive,     // E 欄
      empId,        // F 欄
      status        // G 欄
    ];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    let targetRowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][5]).trim() === empId) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex > 0) {
      sheet.getRange(targetRowIndex, 1, 1, 7).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    // 💡 將完整的 Debug 報告回傳給前端視窗顯示
    return {
      success: true,
      debug_report: {
        received_type: rawInputType,
        received_data: rawInputJson,
        parsed_empId: empId,
        parsed_email: email,
        written_rowData: rowData
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

