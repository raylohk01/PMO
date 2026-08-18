/**
 * 📄 Code.gs
 * Web App 主入口與 API Controller (總控台)
 * 負責處理權限驗證、錯誤攔截與前後端資料交換
 */

// ==========================================
// 1. Web App 入口 (doGet)
// ==========================================
function doGet(e) {
  // 檢查登入狀態 (Google Workspace 預設會強制登入才能訪問)
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    return HtmlService.createHtmlOutput('<h1>存取被拒</h1><p>請先登入 Google 帳號。</p>');
  }

  // 渲染前端頁面
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('HK01 PMO 專案管理系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// 2. 權限驗證輔助函數
// ==========================================
/**
 * 驗證使用者是否具備指定權限
 * @param {Array<string>} allowedRoles - 允許的角色陣列，如 ['Admin', 'PM']
 * @returns {Object} 使用者資訊
 */
function verifyAuth(allowedRoles = []) {
  const email = Session.getActiveUser().getEmail();
  const user = getUserByEmail(email);
  
  if (!user) throw new Error('您尚未在系統中註冊 (Users 表找不到此 Email)。');
  
  if (allowedRoles.length > 0 && !allowedRoles.includes(user.role)) {
    throw new Error(`權限不足！此操作需要以下角色之一：${allowedRoles.join(', ')}`);
  }
  return user;
}

// ==========================================
// 3. API Endpoints (供前端 google.script.run 呼叫)
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

function api_createProject(data) {
  try {
    verifyAuth(['Admin', 'PM', 'Sales']);
    // data 預期包含: clientName, productName, salesPerson, productType
    const newJobNumber = createProject(data.clientName, data.productName, data.salesPerson, data.productType);
    return { success: true, data: { jobNumber: newJobNumber }, message: '專案建立成功' };
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
// [絕對鎖定 textJobType 版 API] 專案整體管理狀態更新
// ==========================================
function api_manageProjectStatus(payloadInput, optionalAction) {
  try {
    let targetJobNumber = '';
    let action = '';

    // 1. 解析前端傳來的參數
    if (typeof payloadInput === 'string' && payloadInput.startsWith('{')) {
      try {
        let parsed = JSON.parse(payloadInput);
        targetJobNumber = String(parsed.jobNumber || parsed.jobNo || parsed.id || '').trim();
        action = String(parsed.action || '').trim();
      } catch(e) {}
    }

    if (!targetJobNumber) {
      targetJobNumber = String(payloadInput || '').trim();
      action = String(optionalAction || '').trim();
    }

    if (!targetJobNumber || targetJobNumber === '[object Object]') {
      throw new Error('未傳入有效的專案編號');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) throw new Error('無專案數據');

    // 2. 尋找欄位 (精準鎖定 textJobType)
    const headersLower = data[0].map(h => String(h || '').trim().toLowerCase());
    
    const idxJobNum = headersLower.findIndex(h => h.includes('jobnumber') || h === 'jobno' || h === 'id');
    const idxStatus = headersLower.findIndex(h => h === 'status' || h === 'project_status');
    
    // 💡 捨棄模糊搜尋，要求完全符合 textjobtype
    const idxTextJob = headersLower.findIndex(h => h === 'textjobtype');

    if (idxJobNum === -1 || idxStatus === -1) throw new Error('工作表缺少 Status 或 JobNumber 欄位');

    let currentUser = 'raylo';
    try {
      const activeEmail = Session.getActiveUser().getEmail();
      if (activeEmail) currentUser = activeEmail.split('@')[0];
    } catch(e){}

    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    let newStatus = '', actionType = '', logDetail = '';

    const actUpper = action.toUpperCase();
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
      throw new Error('未知的操作類型：' + action);
    }

    const newLogEntry = {
      timestamp: nowStr,
      user: currentUser,
      action: actionType,
      details: logDetail
    };

    // 3. 搜尋並更新專案
    for (let i = 1; i < data.length; i++) {
      const currentJob = String(data[i][idxJobNum] || '').trim();

      if (currentJob.toLowerCase() === targetJobNumber.toLowerCase()) {
        
        // A. 更新狀態 (這步確認已經會成功)
        sheet.getRange(i + 1, idxStatus + 1).setValue(newStatus);

        // B. 寫入 textJobType 的 Log
        if (idxTextJob >= 0) {
          let cellValue = String(data[i][idxTextJob] || '').trim();
          let logArray = [];
          
          if (cellValue.startsWith('[')) {
            try {
              logArray = JSON.parse(cellValue);
            } catch(e) {}
          }
          
          if (!Array.isArray(logArray)) {
            logArray = [];
          }
          
          // 將最新 Log 推到最前面
          logArray.unshift(newLogEntry);
          
          // 💡 寫回 Google Sheet 的精準欄位！
          sheet.getRange(i + 1, idxTextJob + 1).setValue(JSON.stringify(logArray));
          
        } else {
          // 防呆警告：如果跑到這裡，代表表格第一列真的沒有叫做 textJobType 的欄位
          throw new Error('專案狀態已更新，但找不到 textJobType 欄位，無法寫入 Log！');
        }

        return {
          success: true,
          message: '專案狀態已更新為 [' + newStatus + ']！'
        };
      }
    }

    throw new Error('找不到編號為 [' + targetJobNumber + '] 的專案');

  } catch (e) {
    return { success: false, message: e.message };
  }
}