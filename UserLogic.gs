// ==========================================
// [UserLogic.gs] 使用者與專案核心邏輯模組 (試算表欄位精準適應與強效防呆版)
// ==========================================

function api_getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    const user = getUserByEmail(email);
    if (!user) {
      return { success: true, data: { email: email, name: email.split('@')[0], department: 'PM', role: 'Management' } };
    }
    return { success: true, data: user };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getUserByEmail(email) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if(!sheet) return null;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idxEmail = headers.findIndex(h => h === 'email');
    const idxName = headers.findIndex(h => h === 'name' || h === 'username');
    const idxDept = headers.findIndex(h => h === 'department' || h === 'team');
    const idxRole = headers.findIndex(h => h === 'role');
    const idxEmpId = headers.findIndex(h => h.includes('empid') || h.includes('emp_id'));
    const idxStatus = headers.findIndex(h => h === 'status');

    if (idxEmail === -1) return null;

    for(let i = 1; i < data.length; i++) {
      if(String(data[i][idxEmail]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
        return {
          empId: idxEmpId >= 0 ? data[i][idxEmpId] : '',
          email: data[i][idxEmail],
          name: idxName >= 0 && data[i][idxName] ? data[i][idxName] : data[i][idxEmail].split('@')[0],
          department: idxDept >= 0 && data[i][idxDept] ? data[i][idxDept] : 'General',
          role: idxRole >= 0 && data[i][idxRole] ? data[i][idxRole] : 'Member',
          status: idxStatus >= 0 ? data[i][idxStatus] : 'Active'
        };
      }
    }
  } catch(e) {
    console.error('getUserByEmail error:', e);
  }
  return null;
}

function api_getAllUsers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if(!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idxEmail = headers.findIndex(h => h === 'email');
    const idxName = headers.findIndex(h => h === 'name' || h === 'username');
    const idxDept = headers.findIndex(h => h === 'department' || h === 'team');
    const idxRole = headers.findIndex(h => h === 'role');
    const idxEmpId = headers.findIndex(h => h.includes('empid') || h.includes('emp_id'));
    const idxStatus = headers.findIndex(h => h === 'status');

    let users = [];
    for(let i = 1; i < data.length; i++) {
      if(data[i][idxEmail]) {
        users.push({
          empId: idxEmpId >= 0 ? data[i][idxEmpId] : '',
          email: data[i][idxEmail],
          name: idxName >= 0 ? data[i][idxName] : '',
          department: idxDept >= 0 ? data[i][idxDept] : '',
          role: idxRole >= 0 ? data[i][idxRole] : '',
          status: idxStatus >= 0 ? data[i][idxStatus] : 'Active'
        });
      }
    }
    return { success: true, data: users };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function api_saveUser(userObj) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxEmpId  = headers.findIndex(h => h.includes('empid') || h.includes('emp_id'));
    const idxEmail  = headers.findIndex(h => h === 'email');
    const idxName   = headers.findIndex(h => h === 'name' || h === 'username');
    const idxDept   = headers.findIndex(h => h === 'department' || h === 'team');
    const idxRole   = headers.findIndex(h => h === 'role');
    const idxActive = headers.findIndex(h => h === 'isactive');
    const idxStatus = headers.findIndex(h => h === 'status');

    const isActive = (userObj.status === 'Active');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (idxEmpId !== -1 && String(data[i][idxEmpId]).trim() === String(userObj.empId).trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      if (idxEmail !== -1)  sheet.getRange(rowIndex, idxEmail + 1).setValue(userObj.email);
      if (idxName !== -1)   sheet.getRange(rowIndex, idxName + 1).setValue(userObj.name);
      if (idxDept !== -1)   sheet.getRange(rowIndex, idxDept + 1).setValue(userObj.department);
      if (idxRole !== -1)   sheet.getRange(rowIndex, idxRole + 1).setValue(userObj.role);
      if (idxActive !== -1) sheet.getRange(rowIndex, idxActive + 1).setValue(isActive);
      if (idxStatus !== -1) sheet.getRange(rowIndex, idxStatus + 1).setValue(userObj.status);
    } else {
      sheet.appendRow([
        userObj.email,
        userObj.name,
        userObj.department,
        userObj.role,
        isActive,
        userObj.empId,
        userObj.status
      ]);
    }
    return { success: true, message: '儲存成功' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function api_getDepartments() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departments');
    if(!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    let depts = [];
    for(let i = 1; i < data.length; i++) {
      if(data[i][0]) {
        depts.push({ name: data[i][0], colorHex: data[i][1] || '#005088' });
      }
    }
    return { success: true, data: depts };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function api_saveDepartment(deptObj) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departments');
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(String(data[i][0]).trim() === String(deptObj.name).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    if(rowIndex > 0) {
      sheet.getRange(rowIndex, 2).setValue(deptObj.colorHex);
    } else {
      sheet.appendRow([deptObj.name, deptObj.colorHex]);
    }
    return { success: true, message: '部門更新成功' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getTemplateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('WorkflowTemplates') || ss.getSheetByName('Templates') || ss.getSheetByName('Template');
}

function api_getTemplates() {
  try {
    const sheet = getTemplateSheet();
    if (!sheet) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    let templates = [];
    for (let i = 1; i < data.length; i++) {
      const nameVal = String(data[i][0] || '').trim();
      const jsonVal = String(data[i][1] || '').trim();
      
      if (nameVal && !nameVal.toLowerCase().includes('templatename') && nameVal.toLowerCase() !== 'name') {
        templates.push({ name: nameVal, json: jsonVal });
      }
    }
    return { success: true, data: templates };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function api_saveTemplate(name, steps) {
  try {
    let sheet = getTemplateSheet();
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('WorkflowTemplates');
      sheet.appendRow(['TemplateName', 'WorkflowJSON']);
    }

    const data = sheet.getDataRange().getValues();
    const jsonStr = JSON.stringify(steps);
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(name).trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 2).setValue(jsonStr);
    } else {
      sheet.appendRow([name, jsonStr]);
    }
    return { success: true, message: '範本儲存成功' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function api_deleteTemplate(name) {
  try {
    const sheet = getTemplateSheet();
    if (!sheet) return { success: true, message: '刪除成功' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(name).trim()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: true, message: '刪除成功' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// [強效保底版] 建立新專案 (含 WorkflowTemplates 範本自動抓取與 PM 關卡自動啟動)
// ==========================================
function api_createProject(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const now = new Date();
    const timeStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");
    const masterPmName = payload.pmName || payload.pm || '未指定';

    // 💡 1. 讀取 WorkflowTemplates 範本庫作為防護備援[cite: 1, 2]
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('WorkflowTemplates');
    let templatesData = templateSheet ? templateSheet.getDataRange().getValues() : [];

    let delivList = payload.deliverables || [];
    
    // 🛡️ 徹底防護：只要發現 deliverables 是空的，強制去 WorkflowTemplates 抓取「標準 Advertorial」範本！[cite: 1]
    if (!Array.isArray(delivList) || delivList.length === 0) {
      var templateName = payload.productType || payload.templateName || '標準 Advertorial';
      var workflowSteps = [];

      if (templateSheet) {
        for (var i = 1; i < templatesData.length; i++) {
          if (String(templatesData[i][0]).trim() === templateName) {
            try { workflowSteps = JSON.parse(templatesData[i][1]); } catch(e){}
            break;
          }
        }
      }

      delivList = [{
        name: payload.deliverableName || '篇章 / 任務 1',
        templateName: templateName,
        templateJson: workflowSteps.length > 0 ? JSON.stringify(workflowSteps) : ''
      }];
    }

    const launchDeadlineStr = payload.deadline || payload.launchDate || Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

    let deliverables = delivList.map((d, idx) => {
      let steps = [];
      let tplName = d.templateName || d.type || payload.productType || '標準 Advertorial';

      if (d.templateJson) {
        try { steps = JSON.parse(d.templateJson); } catch (e) {}
      }

      if (!steps || steps.length === 0) {
        for (let t = 1; t < templatesData.length; t++) {
          if (String(templatesData[t][0]).trim() === String(tplName).trim()) {
            try { steps = JSON.parse(templatesData[t][1]); } catch(e){}
            break;
          }
        }
      }

      // 保底 4 關流程[cite: 2]
      if (!steps || steps.length === 0) {
        steps = [
          { step: 1, name: 'PM 開案與 Briefing', dept: 'PM', fields: ['URL'] },
          { step: 2, name: 'Editor 撰寫', dept: 'Editorial', fields: ['URL'] },
          { step: 3, name: 'Art 設計首圖', dept: 'Design', fields: ['URL'] },
          { step: 4, name: 'Client Review', dept: 'PM', fields: ['URL'] }
        ];
      }

      let formattedSteps = steps.map((s, sIdx) => {
        const dName = s.dept || 'PM';
        const defaultAssignee = (dName.toUpperCase() === 'PM') ? masterPmName : (s.assignee || '');

        return {
          step: sIdx + 1,
          name: s.name || ('步驟 ' + (sIdx + 1)),
          dept: dName,
          assignee: defaultAssignee, // 💡 PM 自動同步指派[cite: 5]
          status: sIdx === 0 ? 'In Progress' : 'Pending', // 💡 關卡 Step 1 自動開啟為 In Progress！[cite: 2, 5]
          isStarted: sIdx === 0 ? true : false,
          fields: s.fields || ['URL'],
          parallelGroup: s.parallelGroup || s.group || '',
          keyDate: launchDeadlineStr,
          deadline: launchDeadlineStr
        };
      });

      return {
        id: 'deliv_' + (idx + 1) + '_' + new Date().getTime(),
        name: d.name || ('篇章/任務 ' + (idx + 1)),
        type: tplName,
        status: 'In Progress', // 💡 任務項目自動設為正式啟動[cite: 2, 5]
        currentStep: 1,
        workflow: formattedSteps
      };
    });

    const workflowData = { deliverables: deliverables };
    const auditLog = [{ timestamp: timeStr, user: masterPmName, action: 'Create & Start Project', details: '建立了此專案並正式啟動專案流程' }];

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxSales = headers.findIndex(h => h.includes('sales'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxLaunch = headers.findIndex(h => h.includes('launch') || h.includes('deadline') || h.includes('死線'));
    const idxSubmission = headers.findIndex(h => h.includes('submission') || h.includes('tentative'));
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxProduct = headers.findIndex(h => h.includes('product') || h.includes('workflow'));
    const idxLog = headers.findIndex(h => h.includes('textjobtype') || h.includes('log') || h.includes('audit'));

    let rowData = new Array(headers.length).fill('');

    if (idxJobNum >= 0) rowData[idxJobNum] = payload.jobNumber || ('A26-' + Math.floor(1000 + Math.random() * 9000));
    if (idxClient >= 0) rowData[idxClient] = payload.client || payload.clientName || '未命名客戶';
    if (idxSales >= 0) rowData[idxSales] = payload.salesName || payload.sales || '未指定';
    if (idxPM >= 0) rowData[idxPM] = masterPmName;
    if (idxSubmission >= 0) rowData[idxSubmission] = payload.tentativeDate || Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
    if (idxLaunch >= 0) rowData[idxLaunch] = payload.deadline || payload.launchDate || Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
    if (idxStatus >= 0) rowData[idxStatus] = 'In Progress'; // 💡 專案狀態預設為 In Progress，免除手動點擊啟動[cite: 2, 5]
    if (idxProduct >= 0) rowData[idxProduct] = JSON.stringify(workflowData);
    if (idxLog >= 0) rowData[idxLog] = JSON.stringify(auditLog);

    sheet.appendRow(rowData);

    return { success: true, message: '專案開案成功並已自動啟動！' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [權限修復版] 我的任務看板資料 (加入 Member 嚴格個人卡片隔離)
// ==========================================
function api_getDashboardData(simEmail) {
  try {
    const userEmail = simEmail || Session.getActiveUser().getEmail();
    const user = getUserByEmail(userEmail);
    const userName = user ? user.name : userEmail.split('@')[0];
    const userRole = user ? user.role : 'Member';
    const userDept = user ? user.department : '';

    // 判斷是否為管理層或部門主管
    const isManager = ['Admin', 'Management', 'Head of PM'].includes(userRole) || (userRole === 'Team Head');

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if(!sheet) return { success: true, data: { overdue:[], dueSoon:[], onTrack:[], upcoming:[], paused:[] } };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: { overdue:[], dueSoon:[], onTrack:[], upcoming:[], paused:[] } };

    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxSales = headers.findIndex(h => h.includes('sales'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxDeadline = headers.findIndex(h => h.includes('launch') || h.includes('deadline') || h.includes('死線') || h.includes('submission'));
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');

    const today = new Date();
    today.setHours(0,0,0,0);

    let result = { overdue: [], dueSoon: [], onTrack: [], upcoming: [], paused: [], actionRequired: [] };

    for(let i = 1; i < data.length; i++) {
      const rawJobNum = idxJobNum >= 0 ? data[i][idxJobNum] : data[i][0];
      const rowJobNum = String(rawJobNum || '').trim();
      if (!rowJobNum || rowJobNum.toLowerCase() === 'jobnumber') continue;

      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim().toLowerCase() : '';
      if (pStatus.includes('recycle') || pStatus.includes('delete') || pStatus.includes('cancel')) continue;

      const pmName = idxPM >= 0 && data[i][idxPM] ? String(data[i][idxPM]).trim() : '未指定';
      const salesName = idxSales >= 0 && data[i][idxSales] ? String(data[i][idxSales]).trim() : '未指定';
      const clientVal = idxClient >= 0 && data[i][idxClient] ? String(data[i][idxClient]).trim() : '未指定客戶';

      let wfData = {};
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables')) {
          try { wfData = JSON.parse(cellStr); break; } catch(e){}
        }
      }

      if (wfData && wfData.deliverables) {
        wfData.deliverables.forEach(d => {
          if (d.status === 'Completed' || d.status === 'Deleted' || d.status === 'Recycle Bin') return;

          const currentStepObj = d.workflow ? d.workflow.find(s => s.step === d.currentStep) : null;
          const currentStepName = currentStepObj ? currentStepObj.name : '未開始';
          const currentAssignee = currentStepObj ? currentStepObj.assignee : '';

          // 💡 權限隔離：一般 Member 在「我的任務」只看得到【指派給自己】或【自己擔當 PM】的卡片！
          if (!isManager) {
            const isMyTask = (currentAssignee && currentAssignee.toLowerCase() === userName.toLowerCase()) || 
                             (pmName && pmName.toLowerCase() === userName.toLowerCase());
            if (!isMyTask) return; // 跳過不屬於該同仁的卡片
          }

          let effectiveDeadlineStr = '';
          if (currentStepObj && (currentStepObj.keyDate || currentStepObj.deadline)) {
            effectiveDeadlineStr = currentStepObj.keyDate || currentStepObj.deadline;
          } else if (d.mainDeadline) {
            effectiveDeadlineStr = d.mainDeadline;
          } else if (idxDeadline >= 0 && data[i][idxDeadline]) {
            let parsed = new Date(data[i][idxDeadline]);
            if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 2000) {
              effectiveDeadlineStr = Utilities.formatDate(parsed, "GMT+8", "yyyy-MM-dd");
            }
          }

          if (!effectiveDeadlineStr) effectiveDeadlineStr = Utilities.formatDate(today, "GMT+8", "yyyy-MM-dd");

          let deadlineDate = new Date(effectiveDeadlineStr);
          const daysLeft = Math.ceil((deadlineDate - today) / (1000 * 60 * 60 * 24));

          const item = {
            jobNumber: rowJobNum,
            client: clientVal,
            taskName: d.name || '未命名任務',
            type: d.type || 'Standard',
            status: d.status || 'Pending Start',
            deadline: effectiveDeadlineStr,
            daysLeft: daysLeft,
            currentStepName: currentStepName,
            assignee: currentAssignee,
            pmName: pmName,
            salesName: salesName
          };

          if (d.status === 'Paused' || pStatus.includes('pause')) {
            result.paused.push(item);
          } else if (d.status === 'Pending Start' || d.status === 'Not Started' || pStatus.includes('pending')) {
            result.upcoming.push(item);
          } else if (daysLeft < 0) {
            result.overdue.push(item);
          } else if (daysLeft <= 3) {
            result.dueSoon.push(item);
          } else {
            result.onTrack.push(item);
          }
        });
      }
    }

    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [分流修復版] 獲取部門營運指揮中心數據 (管線去重與全欄位精準流轉)
// ==========================================
function api_getDeptOperationData(deptName, timeRange, customStart, customEnd) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Projects');
    const usersSheet = ss.getSheetByName('Users');
    
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

    deptName = (deptName || 'Editorial').trim();
    timeRange = timeRange || 'NEXT_14_DAYS';

    let membersMap = {};
    if (usersSheet) {
      const uData = usersSheet.getDataRange().getValues();
      if (uData.length > 1) {
        const headersU = uData[0].map(h => String(h || '').trim().toLowerCase());
        const idxDept = headersU.findIndex(h => h === 'department' || h === 'team' || h === 'dept');
        const idxName = headersU.findIndex(h => h === 'name' || h === 'username');

        for (let i = 1; i < uData.length; i++) {
          const d = idxDept >= 0 ? String(uData[i][idxDept] || '').trim() : '';
          const n = idxName >= 0 ? String(uData[i][idxName] || '').trim() : '';
          if (d.toLowerCase() === deptName.toLowerCase() && n) {
            membersMap[n] = { name: n, inProgress: 0, completed: 0, revisions: 0 };
          }
        }
      }
    }

    let kpi = { overdue: 0, dueSoon: 0, unassigned: 0, onTrack: 0 };
    let calendarMap = {};
    let calendarTasks = [];
    let riskTasks = [];
    let activeTasks = [];
    let pipelineTasks = [];

    const fortyEightHoursLater = new Date(now.getTime() + (48 * 60 * 60 * 1000));

    if (sheet) {
      const data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0].map(h => String(h || '').trim().toLowerCase());
        const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
        const idxClient = headers.findIndex(h => h.includes('client'));
        const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
        const idxDeadline = headers.findIndex(h => h.includes('launch') || h.includes('deadline'));
        const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');

        for (let i = 1; i < data.length; i++) {
          try {
            const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
            if (!jobNumber || jobNumber.toLowerCase() === 'jobnumber') continue;

            const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
            if (pStatus === 'Recycle Bin' || pStatus === 'Cancelled' || pStatus === 'Paused') continue;

            const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '客戶';
            const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';

            let deadlineStr = todayStr;
            if (idxDeadline >= 0 && data[i][idxDeadline]) {
              let parsed = new Date(data[i][idxDeadline]);
              if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 2000) {
                deadlineStr = Utilities.formatDate(parsed, "GMT+8", "yyyy-MM-dd");
              }
            }

            let wfData = {};
            for (let c = 0; c < data[i].length; c++) {
              let cellStr = String(data[i][c] || '');
              if (cellStr.includes('deliverables')) { 
                try { wfData = JSON.parse(cellStr); break; } catch(e){} 
              }
            }

            if (wfData && wfData.deliverables && Array.isArray(wfData.deliverables)) {
              wfData.deliverables.forEach(d => {
                if (!d || d.status === 'Completed') return;

                if (d.workflow && Array.isArray(d.workflow)) {
                  let hasActiveInMyDept = false;

                  // 第一輪：檢查目前部門是否有正在執行的關卡
                  d.workflow.forEach((s) => {
                    const sDept = String(s.dept || '').trim().toLowerCase();
                    const isMyDept = sDept === deptName.toLowerCase();
                    if (isMyDept && (s.status === 'In Progress' || s.step === d.currentStep)) {
                      hasActiveInMyDept = true;
                    }
                  });

                  d.workflow.forEach((s) => {
                    if (!s) return;
                    const sDept = String(s.dept || '').trim().toLowerCase();
                    const targetDept = deptName.toLowerCase();
                    const isMyDept = sDept === targetDept;

                    let assignee = s.assignee || '';
                    if (!assignee && isMyDept && (targetDept === 'pm') && pmName) {
                      assignee = pmName;
                    }

                    // 1. 當前關卡在我們部門（進行中 / 待分配 / 逾期）
                    if (isMyDept && (s.status === 'In Progress' || s.step === d.currentStep)) {
                      if (assignee) {
                        if (!membersMap[assignee]) {
                          membersMap[assignee] = { name: assignee, inProgress: 0, completed: 0, revisions: 0 };
                        }
                        membersMap[assignee].inProgress++;
                      }

                      const effectiveDate = s.keyDate || deadlineStr;
                      calendarMap[effectiveDate] = (calendarMap[effectiveDate] || 0) + 1;

                      const taskObj = {
                        jobNumber: jobNumber,
                        client: clientName,
                        deliverableId: d.id,
                        taskName: d.name || '任務篇章',
                        stepNumber: s.step,
                        stepName: s.name || ('Step ' + s.step),
                        assignee: assignee,
                        deadline: effectiveDate,
                        isOverdue: effectiveDate < todayStr
                      };

                      calendarTasks.push(taskObj);

                      const taskDeadlineDate = new Date(effectiveDate);
                      if (effectiveDate < todayStr) {
                        kpi.overdue++;
                        riskTasks.push(taskObj);
                      } else if (!assignee) {
                        kpi.unassigned++;
                        riskTasks.push(taskObj);
                      } else if (taskDeadlineDate <= fortyEightHoursLater) {
                        kpi.dueSoon++;
                        activeTasks.push(taskObj);
                      } else {
                        kpi.onTrack++;
                        activeTasks.push(taskObj);
                      }
                    }

                    // 2. 未來預備管線：只有當該專案「當前不在我們部門執行」時，後續關卡才納入未來管線！
                    if (isMyDept && s.status === 'Pending' && !hasActiveInMyDept && s.step > d.currentStep) {
                      const currentFront = d.workflow.find(item => item && item.step === d.currentStep);
                      pipelineTasks.push({
                        jobNumber: jobNumber,
                        client: clientName,
                        deliverableId: d.id,
                        taskName: d.name || '任務篇章',
                        stepNumber: s.step,
                        stepName: s.name || ('Step ' + s.step),
                        assignee: assignee || '未指定',
                        frontDept: currentFront ? (currentFront.dept || 'PM') : 'PM',
                        currentFrontStep: currentFront ? ('Step ' + currentFront.step + ' (' + (currentFront.name || '') + ')') : '前置關卡',
                        frontAssignee: currentFront ? (currentFront.assignee || (currentFront.dept + ' 未指派')) : '未指派',
                        estimatedArrival: (d.workflow[0] && d.workflow[0].keyDate) ? d.workflow[0].keyDate : deadlineStr
                      });
                    }

                  });
                }
              });
            }
          } catch(errRow) {
            console.error('Row parse error:', errRow);
          }
        }
      }
    }

    return {
      success: true,
      data: {
        kpi: kpi,
        capacityList: Object.values(membersMap),
        calendarMap: calendarMap,
        calendarTasks: calendarTasks,
        riskTasks: riskTasks,
        activeTasks: activeTasks,
        pipelineTasks: pipelineTasks
      }
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getProjectWorkflow(jobNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxDeadline = headers.findIndex(h => h.includes('launch') || h.includes('deadline'));

    const searchTarget = String(jobNumber).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowJob = String(data[i][idxJobNum >= 0 ? idxJobNum : 0] || '').trim().toLowerCase();
      
      if (rowJob === searchTarget) {
        let deadlineDate = new Date();
        if (idxDeadline >= 0 && data[i][idxDeadline]) {
          let parsed = new Date(data[i][idxDeadline]);
          if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 2000) deadlineDate = parsed;
        }

        let wfCol = -1;
        let wfData = {};
        let logData = [];

        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.includes('deliverables')) { wfCol = c + 1; try { wfData = JSON.parse(cellStr); } catch(e){} }
          if (cellStr.includes('timestamp') && cellStr.includes('action')) { try { logData = JSON.parse(cellStr); } catch(e){} }
        }

        if (!wfData.deliverables || wfData.deliverables.length === 0) {
          wfData = { deliverables: [] };
        }

        let hasChanges = false;
        (wfData.deliverables || []).forEach(d => {
          if (d.workflow) {
            for (let sIdx = 0; sIdx < d.workflow.length; sIdx++) {
              let step = d.workflow[sIdx];
              if (step.status === 'Completed' && step.parallelGroup) {
                let nextStep = d.workflow.find((s, idx) => idx > sIdx && (!s.parallelGroup || s.parallelGroup !== step.parallelGroup) && s.status === 'Pending');
                if (nextStep) {
                  nextStep.status = 'In Progress';
                  if (nextStep.parallelGroup) {
                    d.workflow.filter(s => s.parallelGroup === nextStep.parallelGroup).forEach(s => s.status = 'In Progress');
                  }
                  hasChanges = true;
                }
              }
            }
          }
        });

        if (hasChanges && wfCol > 0) {
          sheet.getRange(i + 1, wfCol).setValue(JSON.stringify(wfData));
        }

        return {
          success: true,
          data: {
            jobNumber: data[i][idxJobNum >= 0 ? idxJobNum : 0],
            client: idxClient >= 0 ? data[i][idxClient] : '未指定客戶',
            pmName: idxPM >= 0 ? data[i][idxPM] : '未指定',
            mainDeadline: Utilities.formatDate(deadlineDate, "GMT+8", "yyyy-MM-dd"),
            workflowData: wfData,
            auditLog: logData
          }
        };
      }
    }
    throw new Error('找不到專案 ' + jobNumber);
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function api_getCompletedProjects(keyword, startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) return { success: true, data: [] };

    const userEmail = Session.getActiveUser().getEmail();
    const currentUser = getUserByEmail(userEmail);

    const isSuperManager = currentUser && (['Admin', 'Management', 'Head of PM'].includes(currentUser.role) || (currentUser.role === 'Team Head' && currentUser.department === 'PM'));
    const isTeamHead = currentUser && currentUser.role === 'Team Head';
    const userName = currentUser ? currentUser.name : userEmail.split('@')[0];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');

    let historyList = [];
    const searchKw = String(keyword || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : String(data[i][0] || '').trim();
      if (!jobNumber || jobNumber.toLowerCase() === 'jobnumber') continue;

      const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '';
      const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';

      let wfData = {};
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables')) {
          try { wfData = JSON.parse(cellStr); break; } catch(e){}
        }
      }

      if (wfData.deliverables) {
        wfData.deliverables.forEach(d => {
          if (d.status === 'Completed') {
            let canView = false;
            if (isSuperManager) {
              canView = true;
            } else if (pmName === userName) {
              canView = true;
            } else if (d.workflow) {
              canView = d.workflow.some(s => (isTeamHead && currentUser && s.dept === currentUser.department) || (s.assignee === userName));
            }

            if (canView) {
              const finishStep = d.workflow ? d.workflow.find(s => s.name.includes('專案完成') || s.name.includes('完成')) : null;
              let launchUrl = '';
              if (finishStep && finishStep.submittedData) {
                launchUrl = Object.values(finishStep.submittedData)[0] || '';
              }

              const matchKw = !searchKw || 
                jobNumber.toLowerCase().includes(searchKw) || 
                clientName.toLowerCase().includes(searchKw) || 
                d.name.toLowerCase().includes(searchKw) ||
                pmName.toLowerCase().includes(searchKw);
              
              if (matchKw) {
                historyList.push({
                  jobNumber: jobNumber,
                  client: clientName,
                  pmName: pmName,
                  deliverableName: d.name,
                  completedAt: finishStep ? finishStep.completedAt : '已完成',
                  launchUrl: launchUrl
                });
              }
            }
          }
        });
      }
    }

    return { success: true, data: historyList };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_updateWorkflowState(jobNumber, deliverableId, payload) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');

    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('找不到該專案');

    let wfCol = -1, logCol = -1;
    let wfData = {}, logData = [];

    for (let c = 0; c < data[rowIndex - 1].length; c++) {
      let cellStr = String(data[rowIndex - 1][c] || '');
      if (cellStr.includes('deliverables')) { wfCol = c + 1; try { wfData = JSON.parse(cellStr); } catch(e){} }
      if (cellStr.includes('timestamp') && cellStr.includes('action')) { logCol = c + 1; try { logData = JSON.parse(cellStr); } catch(e){} }
    }

    if (wfCol === -1) wfCol = 8;
    if (logCol === -1) logCol = 9;

    let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);
    if (!targetD) throw new Error('找不到該子項目');

    const now = new Date();
    const timeStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");

    if (payload.action === 'START') {
      targetD.status = 'In Progress';
      targetD.currentStep = 1;
      if (idxStatus >= 0) sheet.getRange(rowIndex, idxStatus + 1).setValue('In Progress');
      
      const step1 = targetD.workflow.find(s => s.step === 1);
      if (step1) {
        step1.status = 'In Progress';
        if (step1.parallelGroup) {
          targetD.workflow.filter(s => s.parallelGroup && s.parallelGroup === step1.parallelGroup).forEach(s => s.status = 'In Progress');
        }
      }

      logData.unshift({ timestamp: timeStr, user: userName, action: 'Start Deliverable', details: `啟動了任務 [${targetD.name}]` });
    } else if (payload.action === 'SUBMIT') {
      let targetStepNum = payload.stepNumber || targetD.currentStep;
      let currentStepObj = targetD.workflow.find(s => s.step === targetStepNum);
      
      if (currentStepObj) {
        currentStepObj.status = 'Completed';
        currentStepObj.completedAt = timeStr;
        currentStepObj.submittedData = payload.inputs || {};
      }

      let currentGroup = currentStepObj ? currentStepObj.parallelGroup : '';
      
      for (let i = 0; i < targetD.workflow.length; i++) {
        let s = targetD.workflow[i];
        if (s.status === 'Pending') {
          if (!currentGroup || s.parallelGroup !== currentGroup) {
            s.status = 'In Progress';
            if (s.parallelGroup) {
              targetD.workflow.filter(item => item.parallelGroup === s.parallelGroup).forEach(item => item.status = 'In Progress');
            }
            break;
          }
        }
      }

      if (!targetD.workflow.some(s => s.status !== 'Completed')) {
        targetD.status = 'Completed';
      }

      logData.unshift({ timestamp: timeStr, user: userName, action: 'Submit Step', details: `完成了 [${targetD.name}] 的 ${currentStepObj ? currentStepObj.name : ''}` });
    }


// 💡 [關鍵修正] 提交步驟後，自動將 currentStep 更新為當前正處於 In Progress 的步驟
      if (targetD && targetD.workflow) {
        let activeStepObj = targetD.workflow.find(s => s.status === 'In Progress');
        if (activeStepObj) {
          targetD.currentStep = activeStepObj.step;
        } else if (!targetD.workflow.some(s => s.status !== 'Completed')) {
          // 如果全數完成，將 currentStep 設定為最後一步
          targetD.currentStep = targetD.workflow.length;
        }
      }


    sheet.getRange(rowIndex, wfCol).setValue(JSON.stringify(wfData));
    sheet.getRange(rowIndex, logCol).setValue(JSON.stringify(logData));

    return { success: true, message: '更新成功！' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function api_getRecycleBinProjects() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if(!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');

    let recycleList = [];
    for(let i = 1; i < data.length; i++) {
      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
      if(pStatus === 'Recycle Bin') {
        recycleList.push({
          jobNumber: idxJobNum >= 0 ? data[i][idxJobNum] : data[i][0],
          client: idxClient >= 0 ? data[i][idxClient] : '',
          pmName: idxPM >= 0 ? data[i][idxPM] : ''
        });
      }
    }
    return { success: true, data: recycleList };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function api_getPMTeamWorkload() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Users');
    const projectsSheet = ss.getSheetByName('Projects');

    let members = [];
    let stats = {};

    if (usersSheet) {
      const uData = usersSheet.getDataRange().getValues();
      if (uData.length > 1) {
        const headersU = uData[0].map(h => String(h || '').trim().toLowerCase());
        const idxDept = headersU.findIndex(h => h === 'department' || h === 'team' || h === 'dept');
        const idxName = headersU.findIndex(h => h === 'name' || h === 'username');
        const idxStatus = headersU.findIndex(h => h === 'status');

        for (let i = 1; i < uData.length; i++) {
          const d = idxDept >= 0 ? String(uData[i][idxDept] || '').trim().toLowerCase() : '';
          const n = idxName >= 0 ? String(uData[i][idxName] || '').trim() : '';
          const s = idxStatus >= 0 ? String(uData[i][idxStatus] || '').trim() : 'Active';

          if (d === 'pm' && s !== 'Frozen' && n) {
            members.push(n);
            stats[n] = { totalTasks: 0, next5DaysDeliveries: 0, clientCount: 0, clients: new Set() };
          }
        }
      }
    }

    if (projectsSheet && members.length > 0) {
      const pData = projectsSheet.getDataRange().getValues();
      if (pData.length > 1) {
        const headersP = pData[0].map(h => String(h || '').trim().toLowerCase());
        const idxPM = headersP.findIndex(h => h === 'pmname' || h === 'pm');
        const idxStatusP = headersP.findIndex(h => h === 'status' || h === 'project_status');
        const idxClient = headersP.findIndex(h => h === 'clientname' || h === 'client');
        const idxLaunchDate = headersP.findIndex(h => h === 'launchdate' || h.includes('launch') || h.includes('deadline'));

        const now = new Date();
        now.setHours(0,0,0,0);
        const next5DaysEnd = new Date(now.getTime() + (5 * 24 * 60 * 60 * 1000));
        next5DaysEnd.setHours(23,59,59,999);

        for (let i = 1; i < pData.length; i++) {
          const pStatus = idxStatusP >= 0 ? String(pData[i][idxStatusP] || '').trim().toLowerCase() : '';
          if (pStatus === 'completed' || pStatus === 'recycle bin' || pStatus === 'cancelled') continue;

          const pmName = idxPM >= 0 ? String(pData[i][idxPM] || '').trim() : '';
          const clientName = idxClient >= 0 ? String(pData[i][idxClient] || '').trim() : '';
          const rawLaunchDate = idxLaunchDate >= 0 ? pData[i][idxLaunchDate] : null;

          if (pmName) {
            const matchedPM = members.find(m => m.trim().toLowerCase() === pmName.toLowerCase());
            if (matchedPM) {
              stats[matchedPM].totalTasks++;
              if (clientName) stats[matchedPM].clients.add(clientName);

              if (rawLaunchDate) {
                let lDate = new Date(rawLaunchDate);
                if (!isNaN(lDate.getTime())) {
                  lDate.setHours(12,0,0,0);
                  if (lDate >= now && lDate <= next5DaysEnd) {
                    stats[matchedPM].next5DaysDeliveries++;
                  }
                }
              }
            }
          }
        }

        members.forEach(m => {
          stats[m].clientCount = stats[m].clients.size;
          delete stats[m].clients;
        });
      }
    }

    return {
      success: true,
      data: { members: members, stats: stats }
    };

  } catch (e) {
    return { success: false, message: e.message, data: { members: [], stats: {} } };
  }
}

function api_getDeptTeamWorkload(deptName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Users');
    const projectsSheet = ss.getSheetByName('Projects');

    let targetDept = String(deptName || '').trim().toLowerCase();

    let members = [];
    let stats = {};

    if (usersSheet) {
      const uData = usersSheet.getDataRange().getValues();
      if (uData.length > 1) {
        const headersU = uData[0].map(h => String(h || '').trim().toLowerCase());
        const idxDept = headersU.findIndex(h => h === 'department' || h === 'team' || h === 'dept');
        const idxName = headersU.findIndex(h => h === 'name' || h === 'username');
        const idxStatus = headersU.findIndex(h => h === 'status');

        for (let i = 1; i < uData.length; i++) {
          const d = idxDept >= 0 ? String(uData[i][idxDept] || '').trim().toLowerCase() : '';
          const n = idxName >= 0 ? String(uData[i][idxName] || '').trim() : '';
          const s = idxStatus >= 0 ? String(uData[i][idxStatus] || '').trim() : 'Active';

          let isMatch = (d === targetDept) || 
                       (targetDept.includes('edit') && d.includes('edit')) ||
                       (targetDept.includes('design') && d.includes('design')) ||
                       (targetDept.includes('creative') && d.includes('creative')) ||
                       (targetDept.includes('pm') && d.includes('pm'));

          if (isMatch && s !== 'Frozen' && n) {
            members.push(n);
            stats[n] = { totalTasks: 0, next5DaysDeliveries: 0 };
          }
        }
      }
    }

    if (projectsSheet && members.length > 0) {
      const pData = projectsSheet.getDataRange().getValues();
      if (pData.length > 1) {
        const headersP = pData[0].map(h => String(h || '').trim().toLowerCase());
        const idxPM = headersP.findIndex(h => h.includes('pmname') || h === 'pm');
        const idxStatusP = headersP.findIndex(h => h === 'status' || h === 'project_status');
        const idxDeadlineP = headersP.findIndex(h => h.includes('launch') || h.includes('deadline'));

        const now = new Date();
        const next5DaysEnd = new Date(now.getTime() + (5 * 24 * 60 * 60 * 1000));

        for (let i = 1; i < pData.length; i++) {
          const pStatus = idxStatusP >= 0 ? String(pData[i][idxStatusP] || '').trim() : '';
          if (pStatus === 'Completed' || pStatus === 'Recycle Bin' || pStatus === 'Cancelled') continue;

          const pmName = idxPM >= 0 ? String(pData[i][idxPM] || '').trim() : '';

          if (pmName) {
            const matchedPM = members.find(m => m.trim().toLowerCase() === pmName.toLowerCase());
            if (matchedPM) {
              stats[matchedPM].totalTasks++;
              if (idxDeadlineP >= 0 && pData[i][idxDeadlineP]) {
                let pDate = new Date(pData[i][idxDeadlineP]);
                if (!isNaN(pDate.getTime()) && pDate >= now && pDate <= next5DaysEnd) {
                  stats[matchedPM].next5DaysDeliveries++;
                }
              }
            }
          }

          let wfData = null;
          for (let c = 0; c < pData[i].length; c++) {
            let cellStr = String(pData[i][c] || '');
            if (cellStr.includes('deliverables')) {
              try { wfData = JSON.parse(cellStr); break; } catch(e){}
            }
          }

          if (wfData) {
            let deliverablesList = Array.isArray(wfData) ? wfData : (wfData.deliverables || []);
            deliverablesList.forEach(d => {
              if (!d || d.status === 'Completed') return;

              let workflowList = Array.isArray(d.workflow) ? d.workflow : [];
              workflowList.forEach(s => {
                if (!s || s.status === 'Completed') return;

                const assigneeName = String(s.assignee || '').trim();
                if (assigneeName) {
                  const matchedMember = members.find(m => m.trim().toLowerCase() === assigneeName.toLowerCase());
                  if (matchedMember && matchedMember.trim().toLowerCase() !== pmName.toLowerCase()) {
                    stats[matchedMember].totalTasks++;

                    const effectiveDate = s.keyDate || d.mainDeadline || (idxDeadlineP >= 0 ? pData[i][idxDeadlineP] : null);
                    if (effectiveDate) {
                      let kDate = new Date(effectiveDate);
                      if (!isNaN(kDate.getTime()) && kDate >= now && kDate <= next5DaysEnd) {
                        stats[matchedMember].next5DaysDeliveries++;
                      }
                    }
                  }
                }
              });
            });
          }

        }
      }
    }

    return {
      success: true,
      data: { members: members, stats: stats }
    };

  } catch (e) {
    return { success: false, message: e.message, data: { members: [], stats: {} } };
  }
}

// ==========================================
// 💡 [修復 API] 關卡指派並啟動 (若指派 PM 關卡，自動同步右側總 PM 與所有 PM 關卡)
// ==========================================
function api_assignStepAndStart(jobNumber, deliverableId, stepNumber, assignee) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      // 備用搜尋：若未精準傳入 jobNumber，改用 deliverableId 搜尋專案行
      for (let i = 1; i < data.length; i++) {
        for (let c = 0; c < data[i].length; c++) {
          if (String(data[i][c] || '').includes(deliverableId)) {
            rowIndex = i + 1;
            break;
          }
        }
        if (rowIndex > 0) break;
      }
    }

    if (rowIndex === -1) throw new Error('找不到該專案');

    let wfCol = -1, logCol = -1;
    let wfData = {}, logData = [];

    for (let c = 0; c < data[rowIndex - 1].length; c++) {
      let cellStr = String(data[rowIndex - 1][c] || '');
      if (cellStr.includes('deliverables')) { wfCol = c + 1; try { wfData = JSON.parse(cellStr); } catch(e){} }
      if (cellStr.includes('timestamp') && cellStr.includes('action')) { logCol = c + 1; try { logData = JSON.parse(cellStr); } catch(e){} }
    }

    if (wfCol === -1) wfCol = 8;
    if (logCol === -1) logCol = 9;

    let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);
    if (!targetD) throw new Error('找不到該子項目');

    let targetStep = targetD.workflow.find(s => s.step === parseInt(stepNumber));
    if (targetStep) {
      targetStep.assignee = assignee;
      targetStep.isStarted = true;
      targetStep.status = 'In Progress';

      // 💡 核心同步：若變更的是 PM 關卡，自動同步右側總 PM 及同專案所有 PM 關卡 (Step 1 & Step 4)
      if (targetStep.dept && targetStep.dept.toUpperCase() === 'PM') {
        if (idxPM >= 0) {
          sheet.getRange(rowIndex, idxPM + 1).setValue(assignee);
        }
        (targetD.workflow || []).forEach(s => {
          if (s.dept && s.dept.toUpperCase() === 'PM') {
            s.assignee = assignee;
          }
        });
      }

      const now = new Date();
      const timeStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");
      const userEmail = Session.getActiveUser().getEmail();
      const userName = userEmail ? userEmail.split('@')[0] : 'System';

      logData.unshift({
        timestamp: timeStr,
        user: userName,
        action: 'Dispatch Task',
        details: `將 Step ${stepNumber} [${targetStep.name}] 指派給 [${assignee}] 並同步專案 PM`
      });

      sheet.getRange(rowIndex, wfCol).setValue(JSON.stringify(wfData));
      sheet.getRange(rowIndex, logCol).setValue(JSON.stringify(logData));
    }

    return { success: true, message: '指派成功！' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
// 💡 補齊 API：儲存步驟補充資料 / 連結
function api_appendStepData(jobNumber, deliverableId, stepNumber, title, content) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw new Error('找不到專案 ' + jobNumber);

    let wfCol = -1;
    let logCol = -1;
    let wfData = {};
    let logData = [];

    for (let c = 0; c < data[rowIndex - 1].length; c++) {
      let cellStr = String(data[rowIndex - 1][c] || '');
      if (cellStr.includes('deliverables')) { wfCol = c + 1; try { wfData = JSON.parse(cellStr); } catch(e){} }
      if (cellStr.includes('timestamp') && cellStr.includes('action')) { logCol = c + 1; try { logData = JSON.parse(cellStr); } catch(e){} }
    }

    if (wfCol === -1) wfCol = 8;
    if (logCol === -1) logCol = 9;

    let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);
    if (!targetD) throw new Error('找不到該子項目');

    let targetStep = targetD.workflow.find(s => s.step === stepNumber);
    if (targetStep) {
      if (!targetStep.submittedData) targetStep.submittedData = {};
      targetStep.submittedData[title] = content;

      const now = new Date();
      const timeStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");
      const userEmail = Session.getActiveUser().getEmail();
      const userName = userEmail.split('@')[0];

      logData.unshift({
        timestamp: timeStr,
        user: userName,
        action: 'Append Data',
        details: `在 Step ${stepNumber} 補充了資料：[${title}]`
      });

      sheet.getRange(rowIndex, wfCol).setValue(JSON.stringify(wfData));
      sheet.getRange(rowIndex, logCol).setValue(JSON.stringify(logData));
    }

    return { success: true, message: '補充資料儲存成功！' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// 💡 補齊 API：正式啟動客戶審批
function api_startClientReviewStep(jobNumber, deliverableId, stepNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw new Error('找不到專案 ' + jobNumber);

    let wfCol = -1;
    let logCol = -1;
    let wfData = {};
    let logData = [];

    for (let c = 0; c < data[rowIndex - 1].length; c++) {
      let cellStr = String(data[rowIndex - 1][c] || '');
      if (cellStr.includes('deliverables')) { wfCol = c + 1; try { wfData = JSON.parse(cellStr); } catch(e){} }
      if (cellStr.includes('timestamp') && cellStr.includes('action')) { logCol = c + 1; try { logData = JSON.parse(cellStr); } catch(e){} }
    }

    if (wfCol === -1) wfCol = 8;
    if (logCol === -1) logCol = 9;

    let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);
    if (!targetD) throw new Error('找不到該子項目');

    let targetStep = targetD.workflow.find(s => s.step === stepNumber);
    if (targetStep) {
      targetStep.reviewStatus = 'Reviewing';

      const now = new Date();
      const timeStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");
      const userEmail = Session.getActiveUser().getEmail();
      const userName = userEmail.split('@')[0];

      logData.unshift({
        timestamp: timeStr,
        user: userName,
        action: 'Client Review Start',
        details: `將專案送到【客戶追蹤】列中審批`
      });

      sheet.getRange(rowIndex, wfCol).setValue(JSON.stringify(wfData));
      sheet.getRange(rowIndex, logCol).setValue(JSON.stringify(logData));
    }

    return { success: true, message: '已成功送入客戶審批！' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// 💡 補齊 API：動態插入新關卡步驟
function api_insertWorkflowStep(jobNumber, deliverableId, insertAfterStep, newStepObj) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('找不到該專案');

    let wfCol = -1, logCol = -1;
    let wfData = {}, logData = [];

    for (let c = 0; c < data[rowIndex - 1].length; c++) {
      let cellStr = String(data[rowIndex - 1][c] || '');
      if (cellStr.includes('deliverables')) { wfCol = c + 1; try { wfData = JSON.parse(cellStr); } catch(e){} }
      if (cellStr.includes('timestamp') && cellStr.includes('action')) { logCol = c + 1; try { logData = JSON.parse(cellStr); } catch(e){} }
    }
    if (wfCol === -1) wfCol = 8;
    if (logCol === -1) logCol = 9;

    let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);
    if (!targetD) throw new Error('找不到該子項目');

    const insertedStep = {
      step: 0,
      name: newStepObj.name,
      dept: newStepObj.dept,
      status: 'Pending',
      fields: newStepObj.fields || ['URL']
    };

    const targetIdx = parseInt(insertAfterStep);
    targetD.workflow.splice(targetIdx, 0, insertedStep);
    targetD.workflow.forEach((s, idx) => { s.step = idx + 1; });

    const timeStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
    const userName = Session.getActiveUser().getEmail().split('@')[0];

    logData.unshift({
      timestamp: timeStr,
      user: userName,
      action: 'Insert Step',
      details: `動態插入了關卡 Step ${targetIdx + 1}: [${newStepObj.name}] (${newStepObj.dept})`
    });

    sheet.getRange(rowIndex, wfCol).setValue(JSON.stringify(wfData));
    sheet.getRange(rowIndex, logCol).setValue(JSON.stringify(logData));

    return { success: true, message: '成功插入新關卡！' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// [修復版 API] 獲取部門營運指揮中心數據 (支援舊 PM 專案自動承接與真實負載)
// ==========================================
function api_getDeptOperationData(deptName, timeRange, customStart, customEnd) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Projects');
    const usersSheet = ss.getSheetByName('Users');
    
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

    deptName = (deptName || 'Editorial').trim();
    timeRange = timeRange || 'NEXT_14_DAYS';

    let startDate = new Date();
    let endDate = new Date();
    endDate.setDate(now.getDate() + 14);

    if (timeRange === 'TODAY') {
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(23, 59, 59, 999);
    } else if (timeRange === 'LAST_5_DAYS') {
      startDate.setDate(now.getDate() - 5);
      endDate.setHours(23, 59, 59, 999);
    } else if (timeRange === 'LAST_30_DAYS') {
      startDate.setDate(now.getDate() - 30);
      endDate.setHours(23, 59, 59, 999);
    } else if (timeRange === 'CUSTOM' && customStart && customEnd) {
      startDate = new Date(customStart);
      endDate = new Date(customEnd);
      endDate.setHours(23, 59, 59, 999);
    }

    let membersMap = {};
    if (usersSheet) {
      const uData = usersSheet.getDataRange().getValues();
      if (uData.length > 1) {
        const headersU = uData[0].map(h => String(h || '').trim().toLowerCase());
        const idxDept = headersU.findIndex(h => h === 'department' || h === 'team' || h === 'dept');
        const idxName = headersU.findIndex(h => h === 'name' || h === 'username');

        for (let i = 1; i < uData.length; i++) {
          const d = idxDept >= 0 ? String(uData[i][idxDept] || '').trim() : '';
          const n = idxName >= 0 ? String(uData[i][idxName] || '').trim() : '';
          if (d.toLowerCase() === deptName.toLowerCase() && n) {
            membersMap[n] = { name: n, inProgress: 0, completed: 0, revisions: 0 };
          }
        }
      }
    }

    let kpi = { overdue: 0, dueSoon: 0, unassigned: 0, onTrack: 0 };
    let calendarMap = {};
    let calendarTasks = [];
    let riskTasks = [];
    let activeTasks = [];
    let pipelineTasks = [];

    const fortyEightHoursLater = new Date(now.getTime() + (48 * 60 * 60 * 1000));

    if (sheet) {
      const data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0].map(h => String(h || '').trim().toLowerCase());
        const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
        const idxClient = headers.findIndex(h => h.includes('client'));
        const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
        const idxDeadline = headers.findIndex(h => h.includes('launch') || h.includes('deadline'));
        const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');

        for (let i = 1; i < data.length; i++) {
          try {
            const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
            if (!jobNumber || jobNumber.toLowerCase() === 'jobnumber') continue;

            const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
            if (pStatus === 'Recycle Bin' || pStatus === 'Cancelled' || pStatus === 'Paused') continue;

            const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '客戶';
            const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';

            let deadlineStr = todayStr;
            if (idxDeadline >= 0 && data[i][idxDeadline]) {
              let parsed = new Date(data[i][idxDeadline]);
              if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 2000) {
                deadlineStr = Utilities.formatDate(parsed, "GMT+8", "yyyy-MM-dd");
              }
            }

            let wfData = {};
            for (let c = 0; c < data[i].length; c++) {
              let cellStr = String(data[i][c] || '');
              if (cellStr.includes('deliverables')) { 
                try { wfData = JSON.parse(cellStr); break; } catch(e){} 
              }
            }

            if (wfData && wfData.deliverables && Array.isArray(wfData.deliverables)) {
              wfData.deliverables.forEach(d => {
                if (!d || d.status === 'Completed') return;

                if (d.workflow && Array.isArray(d.workflow)) {
                  d.workflow.forEach((s) => {
                    if (!s) return;
                    const sDept = String(s.dept || '').trim().toLowerCase();
                    const targetDept = deptName.toLowerCase();
                    const isMyDept = sDept === targetDept;

                    // 💡 若為 PM 部門且關卡 assignee 為空，自動相容承接專案的 pmName！
                    let assignee = s.assignee || '';
                    if (!assignee && isMyDept && (targetDept === 'pm') && pmName) {
                      assignee = pmName;
                    }

                    if (isMyDept && assignee) {
                      if (!membersMap[assignee]) {
                        membersMap[assignee] = { name: assignee, inProgress: 0, completed: 0, revisions: 0 };
                      }
                      if (s.status === 'Completed') membersMap[assignee].completed++;
                      if (s.revisionCount) membersMap[assignee].revisions += s.revisionCount;
                    }

                    if (isMyDept && (s.status === 'In Progress' || s.status === 'Pending Start' || s.status === 'Pending Assign' || !s.status)) {
                      if (assignee) {
                        if (!membersMap[assignee]) {
                          membersMap[assignee] = { name: assignee, inProgress: 0, completed: 0, revisions: 0 };
                        }
                        membersMap[assignee].inProgress++;
                      }

                      const effectiveDate = s.keyDate || deadlineStr;
                      calendarMap[effectiveDate] = (calendarMap[effectiveDate] || 0) + 1;

                      const taskObj = {
                        jobNumber: jobNumber,
                        client: clientName,
                        deliverableId: d.id,
                        taskName: d.name || '任務篇章',
                        stepNumber: s.step,
                        stepName: s.name || ('Step ' + s.step),
                        assignee: assignee,
                        deadline: effectiveDate,
                        isOverdue: effectiveDate < todayStr
                      };

                      calendarTasks.push(taskObj);

                      const taskDeadlineDate = new Date(effectiveDate);
                      if (effectiveDate < todayStr) {
                        kpi.overdue++;
                        riskTasks.push(taskObj);
                      } else if (!assignee) {
                        kpi.unassigned++;
                        riskTasks.push(taskObj);
                      } else if (taskDeadlineDate <= fortyEightHoursLater) {
                        kpi.dueSoon++;
                        activeTasks.push(taskObj);
                      } else {
                        kpi.onTrack++;
                        activeTasks.push(taskObj);
                      }
                    }

                    if (isMyDept && s.status === 'Pending') {
                      const currentFront = d.workflow.find(item => item && item.status === 'In Progress');
                      if (currentFront) {
                        pipelineTasks.push({
                          jobNumber: jobNumber,
                          client: clientName,
                          deliverableId: d.id,
                          taskName: d.name || '任務篇章',
                          stepNumber: s.step,
                          stepName: s.name || ('Step ' + s.step),
                          frontDept: currentFront.dept || 'PM',
                          currentFrontStep: 'Step ' + currentFront.step + ' (' + (currentFront.name || '') + ')',
                          frontAssignee: currentFront.assignee || ((currentFront.dept || 'PM') + ' 未指派'),
                          estimatedArrival: (d.workflow[0] && d.workflow[0].keyDate) ? d.workflow[0].keyDate : deadlineStr
                        });
                      }
                    }

                  });
                }
              });
            }
          } catch(errRow) {
            console.error('Row parse error:', errRow);
          }
        }
      }
    }

    return {
      success: true,
      data: {
        kpi: kpi,
        capacityList: Object.values(membersMap),
        calendarMap: calendarMap,
        calendarTasks: calendarTasks,
        riskTasks: riskTasks,
        activeTasks: activeTasks,
        pipelineTasks: pipelineTasks
      }
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

// 快捷死線調整回寫 API
function api_updateStepDeadline(deliverableId, stepNumber, newDeadline) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables') && cellStr.includes(deliverableId)) {
          let wfData = JSON.parse(cellStr);
          let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);
          if (targetD && targetD.workflow) {
            let targetS = targetD.workflow.find(s => s.step === parseInt(stepNumber));
            if (targetS) {
              targetS.keyDate = newDeadline;
              sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));
              return { success: true, message: '成功更新死線！' };
            }
          }
        }
      }
    }
    throw new Error('找不到該關卡項目');
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// 💡 補齊 API：快捷派案與組員指派 (若指派的是 PM 關卡，自動雙向同步回寫 Projects 表的 pmName)
function api_dispatchWorkflowStep(jobNumber, deliverableId, stepNumber, assignee, currentDept) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (currentJob.toLowerCase() === String(jobNumber).toLowerCase().trim() || !jobNumber) {

        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.includes('deliverables') && cellStr.includes(deliverableId)) {
            let wfData = JSON.parse(cellStr);
            let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);

            if (targetD && targetD.workflow) {
              let targetS = targetD.workflow.find(s => s.step === parseInt(stepNumber));
              if (targetS) {
                targetS.assignee = assignee;
                targetS.status = 'In Progress';

                // 💡 雙向同步：若被指派的是 PM 關卡，一併回寫更新 Projects 表的 pmName 欄位！
                if (targetS.dept && targetS.dept.toUpperCase() === 'PM' && idxPM >= 0) {
                  sheet.getRange(i + 1, idxPM + 1).setValue(assignee);
                  
                  // 同步該專案所有其他 PM 關卡
                  (targetD.workflow || []).forEach(s => {
                    if (s.dept && s.dept.toUpperCase() === 'PM') s.assignee = assignee;
                  });
                }

                sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));
                return { success: true, message: '派案成功！已指派給 ' + assignee };
              }
            }
          }
        }
      }
    }
    throw new Error('找不到該關卡項目');
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// 💡 補齊 API：修改專案總死線 (Launch Date)
function api_updateProjectDeadline(jobNumber, newDeadline) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxLaunch = headers.findIndex(h => h.includes('launch') || h.includes('deadline') || h.includes('死線'));
    const idxLog = headers.findIndex(h => h.includes('textjobtype') || h.includes('log') || h.includes('audit'));

    if (idxJobNum === -1 || idxLaunch === -1) {
      throw new Error('欄位定位失敗：找不到 JobNumber 或 Launch Date 欄位');
    }

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxJobNum]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw new Error('找不到編號為 [' + jobNumber + '] 的專案');

    sheet.getRange(rowIndex, idxLaunch + 1).setValue(newDeadline);

    if (idxLog >= 0) {
      let logStr = String(data[rowIndex - 1][idxLog] || '').trim();
      let logArray = [];
      if (logStr.startsWith('[')) {
        try { logArray = JSON.parse(logStr); } catch(e) {}
      }
      
      const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
      const userEmail = Session.getActiveUser().getEmail();
      const userName = userEmail ? userEmail.split('@')[0] : 'System';

      logArray.unshift({
        timestamp: nowStr,
        user: userName,
        action: 'Update Deadline',
        details: `將專案總死線修改為：${newDeadline}`
      });

      sheet.getRange(rowIndex, idxLog + 1).setValue(JSON.stringify(logArray));
    }

    return { success: true, message: '專案總死線修改成功！' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getDeptMembers(deptName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Users');
    const projectsSheet = ss.getSheetByName('Projects');

    deptName = String(deptName || 'Editorial').trim().toLowerCase();

    let members = [];
    let memberMap = {};

    if (usersSheet) {
      const uData = usersSheet.getDataRange().getValues();
      if (uData.length > 1) {
        const headersU = uData[0].map(h => String(h || '').trim().toLowerCase());
        const idxDept = headersU.findIndex(h => h === 'department' || h === 'team' || h === 'dept');
        const idxName = headersU.findIndex(h => h === 'name' || h === 'username');
        const idxStatus = headersU.findIndex(h => h === 'status');

        for (let i = 1; i < uData.length; i++) {
          const d = idxDept >= 0 ? String(uData[i][idxDept] || '').trim().toLowerCase() : '';
          const n = idxName >= 0 ? String(uData[i][idxName] || '').trim() : '';
          const s = idxStatus >= 0 ? String(uData[i][idxStatus] || '').trim() : 'Active';

          if (d === deptName && s !== 'Frozen' && n) {
            const memberObj = { name: n, inProgressCount: 0, next5DaysCount: 0 };
            members.push(memberObj);
            memberMap[n.toLowerCase()] = memberObj;
          }
        }
      }
    }

    if (projectsSheet && members.length > 0) {
      const pData = projectsSheet.getDataRange().getValues();
      if (pData.length > 1) {
        const now = new Date();
        const next5DaysEnd = new Date(now.getTime() + (5 * 24 * 60 * 60 * 1000));

        for (let i = 1; i < pData.length; i++) {
          let wfData = {};
          for (let c = 0; c < pData[i].length; c++) {
            let cellStr = String(pData[i][c] || '');
            if (cellStr.includes('deliverables')) {
              try { wfData = JSON.parse(cellStr); break; } catch(e){}
            }
          }

          if (wfData && wfData.deliverables && Array.isArray(wfData.deliverables)) {
            wfData.deliverables.forEach(d => {
              if (!d || d.status === 'Completed') return;

              if (d.workflow && Array.isArray(d.workflow)) {
                d.workflow.forEach(s => {
                  if (!s || !s.assignee) return;

                  const assigneeKey = String(s.assignee || '').trim().toLowerCase();
                  const targetMember = memberMap[assigneeKey];

                  if (targetMember) {
                    if (s.status !== 'Completed') {
                      targetMember.inProgressCount++;

                      if (s.keyDate) {
                        let kDate = new Date(s.keyDate);
                        if (!isNaN(kDate.getTime()) && kDate >= now && kDate <= next5DaysEnd) {
                          targetMember.next5DaysCount++;
                        }
                      }
                    }
                  }
                });
              }
            });
          }
        }
      }
    }

    return {
      success: true,
      members: members
    };

  } catch (e) {
    return { success: false, message: e.message, members: [] };
  }
}

function api_submitWorkflowStep(jobNumber, deliverableId, stepNumber, formData, userEmail) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (currentJob.toLowerCase() === String(jobNumber).toLowerCase().trim()) {

        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.includes('deliverables') && cellStr.includes(deliverableId)) {
            let wfData = JSON.parse(cellStr);
            let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);

            if (targetD && targetD.workflow) {
              const currentStepIdx = targetD.workflow.findIndex(s => s.step === parseInt(stepNumber));
              
              if (currentStepIdx >= 0) {
                const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
                targetD.workflow[currentStepIdx].status = 'Completed';
                targetD.workflow[currentStepIdx].completedAt = nowStr;
                targetD.workflow[currentStepIdx].completedBy = userEmail || 'System';

                if (formData) {
                  targetD.workflow[currentStepIdx].submittedData = formData;
                }

                const nextStepIdx = currentStepIdx + 1;
                if (nextStepIdx < targetD.workflow.length) {
                  targetD.workflow[nextStepIdx].status = 'In Progress';
                } else {
                  targetD.status = 'Completed';
                }

                sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));
                return { success: true, message: '關卡已成功提交並推進至下一階段！' };
              }
            }
          }
        }
      }
    }
    throw new Error('找不到該專案或關卡資料');
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [補齊 API 1] 更換專案總 PM (自動雙向同步所有 PM 關卡)
// ==========================================
function api_reassignPM(jobNumber, newPmName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxLog = headers.findIndex(h => h.includes('textjobtype') || h.includes('log') || h.includes('audit'));

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxJobNum]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw new Error('找不到編號為 [' + jobNumber + '] 的專案');

    // 1. 更新 Projects 工作表欄位中的 pmName
    if (idxPM >= 0) {
      sheet.getRange(rowIndex, idxPM + 1).setValue(newPmName);
    }

    // 2. 雙向同步：走訪 JSON，將所有 dept === 'PM' 的關卡 assignee 一併更新為新 PM
    for (let c = 0; c < data[rowIndex - 1].length; c++) {
      let cellStr = String(data[rowIndex - 1][c] || '');
      if (cellStr.includes('deliverables')) {
        try {
          let wfData = JSON.parse(cellStr);
          (wfData.deliverables || []).forEach(d => {
            (d.workflow || []).forEach(s => {
              if (s.dept && s.dept.toUpperCase() === 'PM') {
                s.assignee = newPmName;
              }
            });
          });
          sheet.getRange(rowIndex, c + 1).setValue(JSON.stringify(wfData));
        } catch(e) {}
        break;
      }
    }

    // 3. 寫入活動日誌 (Log)
    if (idxLog >= 0) {
      let logStr = String(data[rowIndex - 1][idxLog] || '').trim();
      let logArray = [];
      if (logStr.startsWith('[')) {
        try { logArray = JSON.parse(logStr); } catch(e) {}
      }
      
      const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
      const userEmail = Session.getActiveUser().getEmail();
      const userName = userEmail ? userEmail.split('@')[0] : 'System';

      logArray.unshift({
        timestamp: nowStr,
        user: userName,
        action: 'Reassign PM',
        details: `將專案負責 PM 更換為：${newPmName}`
      });

      sheet.getRange(rowIndex, idxLog + 1).setValue(JSON.stringify(logArray));
    }

    return { success: true, message: '專案負責 PM 更換成功！' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [修復 API 2] 關卡指派 (指派 PM 關卡時自動同步總 PM 與 Step 4)
// ==========================================
function api_dispatchWorkflowStep(jobNumber, deliverableId, stepNumber, assignee, currentDept) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (currentJob.toLowerCase() === String(jobNumber).toLowerCase().trim() || !jobNumber) {

        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.includes('deliverables') && cellStr.includes(deliverableId)) {
            let wfData = JSON.parse(cellStr);
            let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);

            if (targetD && targetD.workflow) {
              let targetS = targetD.workflow.find(s => s.step === parseInt(stepNumber));
              if (targetS) {
                targetS.assignee = assignee;
                targetS.status = 'In Progress';

                // 💡 若指派的是 PM 關卡，自動同步回寫 Projects 表的 pmName 及專案內所有 PM 關卡！
                if (targetS.dept && targetS.dept.toUpperCase() === 'PM') {
                  if (idxPM >= 0) sheet.getRange(i + 1, idxPM + 1).setValue(assignee);
                  
                  (targetD.workflow || []).forEach(s => {
                    if (s.dept && s.dept.toUpperCase() === 'PM') s.assignee = assignee;
                  });
                }

                sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));
                return { success: true, message: '派案成功！已指派給 ' + assignee };
              }
            }
          }
        }
      }
    }
    throw new Error('找不到該關卡項目');
  } catch (e) {
    return { success: false, message: e.message };
  }
}


// ==========================================
// 💡 一鍵校正所有專案的 currentStep 欄位
// ==========================================
function fixAllProjectCurrentSteps() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  if(!sheet) return '找不到 Projects 工作表';
  
  const data = sheet.getDataRange().getValues();
  let updatedCount = 0;
  
  for(let i = 1; i < data.length; i++) {
    for (let c = 0; c < data[i].length; c++) {
      let cellStr = String(data[i][c] || '');
      if (cellStr.includes('deliverables')) {
        try {
          let wfData = JSON.parse(cellStr);
          let updated = false;
          
          if (wfData.deliverables) {
            wfData.deliverables.forEach(d => {
              if (d.workflow) {
                let activeStepObj = d.workflow.find(s => s.status === 'In Progress');
                let targetStep = activeStepObj ? activeStepObj.step : (d.workflow.every(s => s.status === 'Completed') ? d.workflow.length : 1);
                
                if (d.currentStep !== targetStep) {
                  d.currentStep = targetStep;
                  updated = true;
                }
              }
            });
          }
          
          if (updated) {
            sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));
            updatedCount++;
          }
        } catch(e){}
      }
    }
  }
  Logger.log(`🎉 校正完成！共修正了 ${updatedCount} 個專案的 currentStep！`);
  return `🎉 校正完成！共修正了 ${updatedCount} 個專案的 currentStep！`;
}