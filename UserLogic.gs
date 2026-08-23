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
// 💡 [Phase B 更新版] 建立新專案 (預設 Pending Start，等待手動啟動)
// ==========================================
function api_createProject(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const now = new Date();
    const timeStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");
    const masterPmName = payload.pmName || payload.pm || '未指定';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('WorkflowTemplates');
    let templatesData = templateSheet ? templateSheet.getDataRange().getValues() : [];

    let delivList = payload.deliverables || [];
    
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
          assignee: defaultAssignee,
          status: 'Pending Start', // 💡 回歸待啟動
          isStarted: false,        // 💡 尚未啟動
          fields: s.fields || ['URL'],
          parallelGroup: s.parallelGroup || s.group || '',
          keyDate: launchDeadlineStr,
          deadline: launchDeadlineStr,
          accumulatedDays: 0      // ⏱️ 初始化累計耗時
        };
      });

      return {
        id: 'deliv_' + (idx + 1) + '_' + new Date().getTime(),
        name: d.name || ('篇章/任務 ' + (idx + 1)),
        type: tplName,
        status: 'Pending Start',  // 💡 任務預設待啟動
        currentStep: 1,
        workflow: formattedSteps
      };
    });

    const workflowData = { deliverables: deliverables };
    const auditLog = [{ timestamp: timeStr, user: masterPmName, action: 'Create Project', details: '建立了此專案 (待啟動)' }];

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
    if (idxStatus >= 0) rowData[idxStatus] = 'Pending Start'; // 💡 預設 Pending Start
    if (idxProduct >= 0) rowData[idxProduct] = JSON.stringify(workflowData);
    if (idxLog >= 0) rowData[idxLog] = JSON.stringify(auditLog);

    sheet.appendRow(rowData);

    return { success: true, message: '專案建立成功！' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [更新版] 我的任務看板數據 (精準區分個人視野與主管部門視野)
// ==========================================
function api_getDashboardData(simEmail) {
  try {
    const userEmail = simEmail || Session.getActiveUser().getEmail();
    const user = getUserByEmail(userEmail);
    const userName = user ? user.name : userEmail.split('@')[0];
    const userRole = user ? user.role : 'Member';
    const userDept = user ? user.department : '';

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

          // 找出真實進行中的關卡
          let activeSteps = d.workflow ? d.workflow.filter(s => s.status === 'In Progress') : [];
          if (activeSteps.length === 0 && d.workflow) {
            let fallback = d.workflow.find(s => s.step === d.currentStep);
            if (fallback) activeSteps.push(fallback);
          }
          const primaryActiveStep = activeSteps.length > 0 ? activeSteps[0] : null;

          let isMyTask = false;
          let isUpcomingForUser = false;

          const isSuperManager = ['Admin', 'Management', 'Head of PM'].includes(userRole);
          const isProjectPM = pmName && pmName.toLowerCase() === userName.toLowerCase();

          if (isSuperManager || isProjectPM) {
            isMyTask = true;
          } else {
            let activeMatch = false;
            let futureMatch = false;

            // 💡 關鍵修正：依照身份切換「關心範圍」
            if (userRole === 'Team Head') {
              // 主管視野：只要關卡屬於我的「部門」，我就在乎！
              activeMatch = activeSteps.some(s => s.dept === userDept);
              futureMatch = d.workflow && d.workflow.some(s => s.step > d.currentStep && s.dept === userDept);
            } else {
              // 組員視野：只有明確「指派給我」的關卡，我才在乎！
              activeMatch = activeSteps.some(s => s.assignee && s.assignee.toLowerCase() === userName.toLowerCase());
              futureMatch = d.workflow && d.workflow.some(s => s.step > d.currentStep && s.assignee && s.assignee.toLowerCase() === userName.toLowerCase());
            }

            if (activeMatch) {
              isMyTask = true;
            } else if (futureMatch) {
              isMyTask = true;
              isUpcomingForUser = true; // 因為目前還沒輪到我(或我的部門)
            }
          }

          if (!isMyTask) return; // 與自己或部門無關，跳過

          const displayStep = primaryActiveStep;
          const currentStepName = displayStep ? (displayStep.dept + ' ' + displayStep.name) : '未開始';
          const currentAssignee = displayStep ? displayStep.assignee : '';

          let effectiveDeadlineStr = '';
          if (displayStep && (displayStep.keyDate || displayStep.deadline)) {
            effectiveDeadlineStr = displayStep.keyDate || displayStep.deadline;
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
            jobNumber: rowJobNum, client: clientVal, taskName: d.name || '未命名任務',
            type: d.type || 'Standard', status: d.status || 'Pending Start',
            deadline: effectiveDeadlineStr, daysLeft: daysLeft,
            currentStepName: currentStepName, assignee: currentAssignee,
            pmName: pmName, salesName: salesName
          };

          let isPaused = d.status === 'Paused' || pStatus.includes('pause');
          let isActive = primaryActiveStep && primaryActiveStep.status === 'In Progress';

          if (isPaused) {
            result.paused.push(item);
          } else if (isUpcomingForUser) {
            result.upcoming.push(item); 
          } else if (isActive) {
            if (daysLeft < 0) result.overdue.push(item);
            else if (daysLeft <= 3) result.dueSoon.push(item);
            else result.onTrack.push(item);
          } else {
            result.upcoming.push(item);
          }
        });
      }
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/*

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

*/

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

// ==========================================
// 💡 [更新版] 獲取已完成專案歷史 (支援提交日、死線、遲交結算與精準 URL 抓取)
// ==========================================
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
    const idxSubmission = headers.findIndex(h => h.includes('submission') || h.includes('tentative') || h.includes('created'));
    const idxDeadline = headers.findIndex(h => h.includes('launch') || h.includes('deadline') || h.includes('死線'));

    let historyList = [];
    const searchKw = String(keyword || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : String(data[i][0] || '').trim();
      if (!jobNumber || jobNumber.toLowerCase() === 'jobnumber') continue;

      const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '';
      const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';

      let submissionStr = (idxSubmission >= 0 && data[i][idxSubmission]) ? Utilities.formatDate(new Date(data[i][idxSubmission]), "GMT+8", "yyyy-MM-dd") : '未設定';
      let deadlineStr = (idxDeadline >= 0 && data[i][idxDeadline]) ? Utilities.formatDate(new Date(data[i][idxDeadline]), "GMT+8", "yyyy-MM-dd") : '未設定';

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
            if (isSuperManager) canView = true;
            else if (pmName === userName) canView = true;
            else if (d.workflow) canView = d.workflow.some(s => (isTeamHead && currentUser && s.dept === currentUser.department) || (s.assignee === userName));

            if (canView) {
              // 1. 抓取精確完成時間
              let compAt = d.completedAt || '已完成';
              let compAtIso = d.completedAtIso || null;
              
              if (!compAtIso && d.workflow) {
                const lastStep = d.workflow[d.workflow.length - 1];
                if (lastStep && lastStep.completedAtIso) {
                  compAtIso = lastStep.completedAtIso;
                  compAt = lastStep.completedAt;
                }
              }

              // 2. 計算是否遲交
              let delayText = '';
              let isDelayed = false;
              if (compAtIso && deadlineStr !== '未設定') {
                let deadlineDate = new Date(deadlineStr + 'T23:59:59'); // 以死線當天晚上 23:59 為界線
                let compDate = new Date(compAtIso);
                let diffMs = compDate.getTime() - deadlineDate.getTime();
                
                if (diffMs > 0) {
                  isDelayed = true;
                  let dDays = Math.floor(diffMs / 86400000);
                  let dHrs = Math.floor((diffMs % 86400000) / 3600000);
                  let dMins = Math.floor((diffMs % 3600000) / 60000);
                  delayText = `遲了 ${dDays}天 ${dHrs}小時 ${dMins}分`;
                } else {
                  delayText = '按時完成';
                }
              } else {
                delayText = '無死線紀錄';
              }

              // 💡 3. 強效抓取上線連結 (倒序掃描所有關卡，抓取最後一個填寫的網址)
              let launchUrl = '';
              if (d.workflow) {
                for (let sIdx = d.workflow.length - 1; sIdx >= 0; sIdx--) {
                  let sData = d.workflow[sIdx].submittedData;
                  if (sData) {
                    for (let key in sData) {
                      let val = String(sData[key]).trim();
                      if (val.includes('http') || val.includes('www') || val.includes('.com')) {
                        launchUrl = val;
                        break;
                      }
                    }
                  }
                  if (launchUrl) break;
                }
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
                  submissionDate: submissionStr,
                  deadline: deadlineStr,
                  completedAt: compAt,
                  completedAtIso: compAtIso || '1970-01-01T00:00:00.000Z',
                  delayText: delayText,
                  isDelayed: isDelayed,
                  launchUrl: launchUrl
                });
              }
            }
          }
        });
      }
    }

    // 💡 4. 依照完成時間由新到舊排序 (Descending)
    historyList.sort((a, b) => new Date(b.completedAtIso).getTime() - new Date(a.completedAtIso).getTime());

    return { success: true, data: historyList };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [更新版] 推進工作流 (支援 Client Review 等待送出計時)
// ==========================================
function api_updateWorkflowState(jobNumber, deliverableId, payload) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxLog = headers.findIndex(h => h.includes('textjobtype') || h.includes('log') || h.includes('audit'));

    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('找不到該專案');

    let wfCol = -1, wfData = {};
    for (let c = 0; c < data[rowIndex - 1].length; c++) {
      let cellStr = String(data[rowIndex - 1][c] || '');
      if (cellStr.includes('deliverables')) { wfCol = c + 1; try { wfData = JSON.parse(cellStr); } catch(e){} }
    }
    if (wfCol === -1) wfCol = 8;

    let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);
    if (!targetD) throw new Error('找不到該子項目');

    const now = new Date();
    const timeStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");
    const isoNowStr = now.toISOString();

    if (payload.action === 'START') {
      targetD.status = 'In Progress';
      targetD.currentStep = 1;
      targetD.startedAt = isoNowStr;
      if (idxStatus >= 0) sheet.getRange(rowIndex, idxStatus + 1).setValue('In Progress');
      
      const step1 = targetD.workflow.find(s => s.step === 1);
      if (step1) {
        step1.status = 'In Progress';
        step1.isStarted = true;
        step1.startedAt = isoNowStr; 
        step1.dispatchWaitMs = 0;
      }
      writeTextJobTypeLog(sheet, rowIndex, idxLog, timeStr, userName, 'Start Project', `正式啟動了任務 [${targetD.name}]`);
    } 
    else if (payload.action === 'SUBMIT') {
      let targetStepNum = payload.stepNumber || targetD.currentStep;
      let currentStepObj = targetD.workflow.find(s => s.step === targetStepNum);
      
      if (currentStepObj) {
        currentStepObj.status = 'Completed';
        currentStepObj.completedAt = timeStr;
        currentStepObj.completedAtIso = isoNowStr;
        currentStepObj.submittedData = payload.inputs || {};

        if (currentStepObj.startedAt) {
          let spentMs = Math.max(0, now.getTime() - new Date(currentStepObj.startedAt).getTime());
          currentStepObj.accumulatedMs = (currentStepObj.accumulatedMs || 0) + spentMs;
        }
      }

      let currentGroup = currentStepObj ? currentStepObj.parallelGroup : '';
      
      for (let i = 0; i < targetD.workflow.length; i++) {
        let s = targetD.workflow[i];
        if (s.status === 'Pending' || s.status === 'Pending Start') {
          if (!currentGroup || s.parallelGroup !== currentGroup) {
            s.status = 'In Progress';
            s.pendingAssignmentAt = isoNowStr; 

            // 💡 特例處理：如果是 Client Review 相關關卡，即使分配給了 PM，馬錶也不要啟動！
            const isClientStep = String(s.dept || '').toLowerCase().includes('client') || 
                                 String(s.name || '').toLowerCase().includes('client') || 
                                 String(s.name || '').toLowerCase().includes('review');

            if (s.assignee && !isClientStep) {
              s.startedAt = isoNowStr;
              s.dispatchWaitMs = 0;
            } else {
              s.startedAt = null; // 尚未派案，或等待 PM 送出審批，暫不計時
            }

            if (s.parallelGroup) {
              targetD.workflow.filter(item => item.parallelGroup === s.parallelGroup).forEach(item => {
                item.status = 'In Progress';
                item.pendingAssignmentAt = isoNowStr;
                const isCStep = String(item.dept || '').toLowerCase().includes('client') || 
                                String(item.name || '').toLowerCase().includes('client') || 
                                String(item.name || '').toLowerCase().includes('review');
                if (item.assignee && !isCStep) {
                  item.startedAt = isoNowStr;
                  item.dispatchWaitMs = 0;
                } else {
                  item.startedAt = null;
                }
              });
            }
            break;
          }
        }
      }

      if (!targetD.workflow.some(s => s.status !== 'Completed')) {
        targetD.status = 'Completed';
        targetD.completedAt = timeStr;
        targetD.completedAtIso = isoNowStr;
      }

      writeTextJobTypeLog(sheet, rowIndex, idxLog, timeStr, userName, 'Submit Step', `完成了 [${targetD.name}] 的 ${currentStepObj ? currentStepObj.name : ''}`);
    }

    if (targetD && targetD.workflow) {
      let activeStepObj = targetD.workflow.find(s => s.status === 'In Progress');
      if (activeStepObj) targetD.currentStep = activeStepObj.step;
      else if (!targetD.workflow.some(s => s.status !== 'Completed')) targetD.currentStep = targetD.workflow.length;
    }

    sheet.getRange(rowIndex, wfCol).setValue(JSON.stringify(wfData));
    return { success: true, message: '更新成功！' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// 💡 寫入 textJobType Log 的專用輔助函數
function writeTextJobTypeLog(sheet, rowIndex, idxLog, timeStr, user, action, details) {
  if (idxLog < 0) return;
  let logStr = String(sheet.getRange(rowIndex, idxLog + 1).getValue() || '').trim();
  let logArray = [];
  if (logStr.startsWith('[')) {
    try { logArray = JSON.parse(logStr); } catch(e) {}
  }
  logArray.unshift({ timestamp: timeStr, user: user, action: action, details: details });
  sheet.getRange(rowIndex, idxLog + 1).setValue(JSON.stringify(logArray));
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

// ==========================================
// 💡 [更新版] 正式啟動客戶審批 (結算準備時間，啟動審批馬錶)
// ==========================================
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

    let wfCol = -1, logCol = -1, wfData = {}, logData = [];

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
      const nowIso = now.toISOString();
      const timeStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");
      
      // 💡 [核心算式] 結算「等待 PM 處理/準備的時間」，並正式開啟客戶審批的馬錶！
      if (!targetStep.startedAt) {
        targetStep.startedAt = nowIso;
        if (targetStep.pendingAssignmentAt) {
          let waitMs = Math.max(0, now.getTime() - new Date(targetStep.pendingAssignmentAt).getTime());
          targetStep.dispatchWaitMs = waitMs;
        } else {
          targetStep.dispatchWaitMs = 0;
        }
      }

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
// 💡 部門營運數據 API (顯示真實關卡現狀版)
// ==========================================
function api_getDeptOperationData(dept, timeRange, startDate, endDate) {
  try {
    const targetDept = String(dept || 'Editorial').trim();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');

    let capacityMap = {};
    let calendarMap = {};
    let calendarTasks = [];
    
    let riskTasks = [];
    let activeTasks = [];
    let pipelineTasks = [];
    let clientReviewTasks = [];
    let completedTasks = [];

    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

    for (let i = 1; i < data.length; i++) {
      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
      if (pStatus === 'Completed' || pStatus === 'Recycle Bin' || pStatus === 'Cancelled') continue;

      let wfData = {};
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables')) {
          try { wfData = JSON.parse(cellStr); break; } catch(e){}
        }
      }

      if (wfData && wfData.deliverables) {
        wfData.deliverables.forEach(d => {
          if (d.status === 'Deleted') return;

          const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '') : '';
          const client = idxClient >= 0 ? String(data[i][idxClient] || '') : '';
          const pmName = idxPM >= 0 ? String(data[i][idxPM] || '') : '';

          if (!d.workflow || d.workflow.length === 0) return;

          // 1. 組員產能統計
          d.workflow.forEach(s => {
            const sDept = String(s.dept || '').trim().toLowerCase();
            const assignee = String(s.assignee || '').trim();
            if (sDept === targetDept.toLowerCase() && assignee) {
              if (!capacityMap[assignee]) {
                capacityMap[assignee] = { name: assignee, inProgress: 0, completed: 0, revisions: 0 };
              }
              if (s.status === 'In Progress') capacityMap[assignee].inProgress++;
              if (s.status === 'Completed') capacityMap[assignee].completed++;
            }
          });

          // 2. 互斥歸類與動態卡片封面標籤
          const activeStep = d.workflow.find(s => s.status === 'In Progress');
          const allCompleted = d.status === 'Completed' || d.workflow.every(s => s.status === 'Completed');
          const deptCompletedSteps = d.workflow.filter(s => String(s.dept || '').trim().toLowerCase() === targetDept.toLowerCase() && s.status === 'Completed');

          if (allCompleted) {
            // 全數完成 -> 封面關卡標示為「已完成」
            if (deptCompletedSteps.length > 0) {
              const lastStep = deptCompletedSteps[deptCompletedSteps.length - 1];
              let compDateStr = lastStep.completedAt ? lastStep.completedAt.split(' ')[0] : todayStr;
              let compDate = new Date(compDateStr);
              let diffDays = Math.floor((now - compDate) / (1000 * 60 * 60 * 24));

              if (diffDays <= 5) {
                const compTask = {
                  jobNumber: jobNumber,
                  client: client,
                  taskName: d.name || '未命名任務',
                  stepNumber: lastStep.step,
                  stepName: '已完成', // 💡 整體皆已完成，封面明確印出「已完成」
                  assignee: lastStep.assignee || pmName,
                  deliverableId: d.id,
                  deadline: compDateStr,
                  category: 'COMPLETED'
                };
                completedTasks.push(compTask);
                calendarTasks.push(compTask);
                calendarMap[compDateStr] = (calendarMap[compDateStr] || 0) + 1;
              }
            }
          } else if (activeStep) {
            const activeDept = String(activeStep.dept || '').trim().toLowerCase();
            const assignee = String(activeStep.assignee || '').trim();
            const deadline = activeStep.deadline || activeStep.keyDate || todayStr;
            const reviewStatus = String(activeStep.reviewStatus || '').trim();
            const currentRealStepName = activeStep.name || ('Step ' + activeStep.step);

            if (activeDept === targetDept.toLowerCase()) {
              const isClientReview = (reviewStatus === 'Reviewing') || 
                                     (currentRealStepName.toLowerCase().includes('client')) || 
                                     (activeDept.includes('client'));

              let taskCategory = 'ACTIVE';
              if (isClientReview) taskCategory = 'CLIENT_REVIEW';
              else if (deadline < todayStr) taskCategory = 'OVERDUE';
              else if (!assignee) taskCategory = 'UNASSIGNED';
              else {
                let daysLeft = Math.ceil((new Date(deadline) - new Date(todayStr)) / 86400000);
                if (daysLeft <= 2) taskCategory = 'DUE_SOON';
              }

              const taskObj = {
                jobNumber: jobNumber,
                client: client,
                taskName: d.name || '未命名任務',
                stepNumber: activeStep.step,
                stepName: currentRealStepName,
                assignee: assignee || pmName,
                deadline: deadline,
                isOverdue: (deadline < todayStr),
                deliverableId: d.id,
                pmName: pmName,
                category: taskCategory
              };

              if (deadline) {
                calendarTasks.push(taskObj);
                calendarMap[deadline] = (calendarMap[deadline] || 0) + 1;
              }

              if (taskCategory === 'CLIENT_REVIEW') clientReviewTasks.push(taskObj);
              else if (taskCategory === 'OVERDUE' || taskCategory === 'UNASSIGNED') riskTasks.push(taskObj);
              else activeTasks.push(taskObj);

            } else {
              // 當前關卡在其他部門
              const futureStep = d.workflow.find(s => s.step > activeStep.step && String(s.dept || '').trim().toLowerCase() === targetDept.toLowerCase());
              if (futureStep) {
                pipelineTasks.push({
                  jobNumber: jobNumber,
                  client: client,
                  taskName: d.name || '未命名任務',
                  stepName: futureStep.name,
                  currentFrontStep: currentRealStepName + ' (' + activeStep.dept + ')',
                  pmName: pmName,
                  deliverableId: d.id
                });
              } else if (deptCompletedSteps.length > 0) {
                // 本部門已完成，但專案整體卡在下游關卡 -> 封面顯示真實當前關卡 (如 Client Review)
                const lastStep = deptCompletedSteps[deptCompletedSteps.length - 1];
                let compDateStr = lastStep.completedAt ? lastStep.completedAt.split(' ')[0] : todayStr;
                let compDate = new Date(compDateStr);
                let diffDays = Math.floor((now - compDate) / (1000 * 60 * 60 * 24));

                if (diffDays <= 5) {
                  const compTask = {
                    jobNumber: jobNumber,
                    client: client,
                    taskName: d.name || '未命名任務',
                    stepNumber: lastStep.step,
                    stepName: currentRealStepName, // 💡 印出專案當前真實關卡 (如 Client Review)
                    assignee: lastStep.assignee || pmName,
                    deliverableId: d.id,
                    deadline: compDateStr,
                    category: 'COMPLETED'
                  };
                  completedTasks.push(compTask);
                  calendarTasks.push(compTask);
                  calendarMap[compDateStr] = (calendarMap[compDateStr] || 0) + 1;
                }
              }
            }
          }
        });
      }
    }

    const capacityList = Object.keys(capacityMap).map(k => capacityMap[k]);

    return {
      success: true,
      data: {
        capacityList: capacityList,
        calendarMap: calendarMap,
        calendarTasks: calendarTasks,
        riskTasks: riskTasks,
        activeTasks: activeTasks,
        pipelineTasks: pipelineTasks,
        clientReviewTasks: clientReviewTasks,
        completedTasks: completedTasks
      }
    };
  } catch (e) {
    return { success: false, message: 'api_getDeptOperationData 錯誤: ' + e.message };
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

// ==========================================
// 💡 [更新版] 派案 API (精準計算「派案等待時間」並啟動工作馬錶)
// ==========================================
function api_dispatchWorkflowStep(jobNumber, deliverableId, stepNumber, assignee, currentDept) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('textjobtype') || h.includes('audit'));

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
                const nowIso = new Date().toISOString();

                // 💡 [核心算式] 若關卡已是 In Progress 且尚未開始動工：正式啟動工作計時，並計算派案等待時間！
                if (targetS.status === 'In Progress' && !targetS.startedAt) {
                  targetS.startedAt = nowIso;
                  
                  if (targetS.pendingAssignmentAt) {
                    let waitMs = Math.max(0, new Date(nowIso).getTime() - new Date(targetS.pendingAssignmentAt).getTime());
                    targetS.dispatchWaitMs = waitMs; // ⏱️ 精確記錄「等待主管派案」的時間差
                  } else {
                    targetS.dispatchWaitMs = 0;
                  }
                }

                if (targetS.dept && targetS.dept.toUpperCase() === 'PM' && idxPM >= 0) {
                  sheet.getRange(i + 1, idxPM + 1).setValue(assignee);
                  (targetD.workflow || []).forEach(s => {
                    if (s.dept && s.dept.toUpperCase() === 'PM') s.assignee = assignee;
                  });
                }

                sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));

                if (idxAudit >= 0) {
                  let logs = [];
                  let cellStrLog = String(data[i][idxAudit] || '').trim();
                  if (cellStrLog.startsWith('[')) {
                    try { logs = JSON.parse(cellStrLog); } catch(e) {}
                  } else if (cellStrLog) {
                    logs = [{ timestamp: "系統紀錄", user: "System", action: cellStrLog }];
                  }
                  const activeUser = Session.getActiveUser().getEmail().split('@')[0];
                  const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
                  logs.unshift({
                    timestamp: nowStr,
                    user: activeUser,
                    action: `將 Step ${targetS.step} [${targetS.name}] 指派給 [${assignee}]`
                  });
                  sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
                }

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

// ==========================================
// ⏱️ Phase B: 專案耗時與工作天計算機 (扣除週六日，精確至小數點)
// ==========================================
function calculateWorkingDays(startDateStr, endDateStr) {
  if (!startDateStr || !endDateStr) return 0;
  
  let start = new Date(startDateStr);
  let end = new Date(endDateStr);
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || start > end) return 0;

  let diffMs = end.getTime() - start.getTime();
  let totalDays = diffMs / (1000 * 60 * 60 * 24); 

  let weekendDays = 0;
  let cur = new Date(start);
  cur.setHours(0,0,0,0);
  let endMidnight = new Date(end);
  endMidnight.setHours(0,0,0,0);

  while (cur <= endMidnight) {
    let day = cur.getDay();
    if (day === 0 || day === 6) weekendDays++; 
    cur.setDate(cur.getDate() + 1);
  }

  let workingDays = totalDays - weekendDays;
  return Math.max(0, Number(workingDays.toFixed(2))); 
}

// ==========================================
// 💡 [補齊 API] 客戶退回修改 (Client Revision) - 支援 Phase B 計時累加
// ==========================================
function api_triggerClientRevision(jobNumber, deliverableId, stepNumber, feedback) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('textjobtype') || h.includes('audit'));

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (currentJob.toLowerCase() === String(jobNumber).toLowerCase().trim() || !jobNumber) {

        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.includes('deliverables') && cellStr.includes(deliverableId)) {
            let wfData = JSON.parse(cellStr);
            let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);

            if (targetD && targetD.workflow) {
              // 找到要被「退回重做」的目標關卡
              let targetS = targetD.workflow.find(s => s.step === parseInt(stepNumber));
              if (targetS) {
                
                const nowIso = new Date().toISOString();

                // 1. 將目標關卡改回 In Progress 並重新啟動馬錶
                targetS.status = 'In Progress';
                targetS.completedAt = null;
                targetS.completedAtIso = null;
                targetS.remarks = `【客戶退回修改】${feedback}`; // 顯示退回理由
                
                // 💡 雙軌計時：如果有關卡負責人，立刻開始計時；沒有則進入等待派案
                if (targetS.assignee) {
                  targetS.startedAt = nowIso; 
                  targetS.dispatchWaitMs = 0;
                } else {
                  targetS.startedAt = null;
                  targetS.pendingAssignmentAt = nowIso;
                }
                
                // 2. 將目標關卡「之後」一直到「現在」的所有關卡，全部重置為 Pending
                targetD.workflow.forEach(s => {
                  if (s.step > targetS.step && s.status !== 'Pending Start') {
                    s.status = 'Pending';
                    s.startedAt = null;
                    s.pendingAssignmentAt = null;
                    s.completedAt = null;
                    s.completedAtIso = null;
                    s.reviewStatus = null; // 清除審批狀態
                    s.remarks = ''; // 清除後續關卡的警告
                    // 💡 注意：我們保留了 s.accumulatedMs (累計耗時)，確保之前的努力不會白費
                  }
                });

                // 3. 調整當前步驟指標
                targetD.currentStep = targetS.step;
                targetD.status = 'In Progress';

                sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));

                // 4. 寫入活動日誌 (Log)
                if (idxAudit >= 0) {
                  let logs = [];
                  let cellStrLog = String(data[i][idxAudit] || '').trim();
                  if (cellStrLog.startsWith('[')) {
                    try { logs = JSON.parse(cellStrLog); } catch(e) {}
                  } else if (cellStrLog) {
                    logs = [{ timestamp: "系統紀錄", user: "System", action: cellStrLog }];
                  }
                  
                  const activeUser = Session.getActiveUser().getEmail().split('@')[0];
                  const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
                  
                  logs.unshift({
                    timestamp: nowStr,
                    user: activeUser,
                    action: `退回修改 (Revision)`,
                    details: `客戶退回至 Step ${targetS.step} [${targetS.name}]。原因：${feedback}`
                  });
                  sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
                }

                return { success: true, message: '成功退回修改！' };
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
// 💡 Phase C: 關卡強制退回上一步 (Rollback)
// ==========================================
function api_rollbackWorkflowStep(jobNumber, deliverableId, targetStepNumber, reason) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('textjobtype') || h.includes('audit'));

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (currentJob.toLowerCase() === String(jobNumber).toLowerCase().trim() || !jobNumber) {

        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.includes('deliverables') && cellStr.includes(deliverableId)) {
            let wfData = JSON.parse(cellStr);
            let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);

            if (targetD && targetD.workflow) {
              let rollbackStepNum = parseInt(targetStepNumber);
              let targetS = targetD.workflow.find(s => s.step === rollbackStepNum);
              if (!targetS) throw new Error('找不到目標退回關卡');

              const nowIso = new Date().toISOString();

              // 1. 將目標關卡設回 In Progress
              targetS.status = 'In Progress';
              targetS.completedAt = null;
              targetS.completedAtIso = null;
              targetS.remarks = `【PM 強制退回】${reason}`;

              if (targetS.assignee) {
                targetS.startedAt = nowIso;
                targetS.dispatchWaitMs = 0;
              } else {
                targetS.startedAt = null;
                targetS.pendingAssignmentAt = nowIso;
              }

              // 2. 將目標關卡之後的所有關卡重置為 Pending
              targetD.workflow.forEach(s => {
                if (s.step > rollbackStepNum) {
                  s.status = 'Pending';
                  s.startedAt = null;
                  s.pendingAssignmentAt = null;
                  s.completedAt = null;
                  s.completedAtIso = null;
                  s.reviewStatus = null;
                  s.remarks = '';
                }
              });

              targetD.currentStep = rollbackStepNum;
              targetD.status = 'In Progress';

              if (idxStatus >= 0) sheet.getRange(i + 1, idxStatus + 1).setValue('In Progress');
              sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));

              // 3. 寫入活動日誌
              if (idxAudit >= 0) {
                let logs = [];
                let cellStrLog = String(data[i][idxAudit] || '').trim();
                if (cellStrLog.startsWith('[')) {
                  try { logs = JSON.parse(cellStrLog); } catch(e) {}
                }
                const activeUser = Session.getActiveUser().getEmail().split('@')[0];
                const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
                
                logs.unshift({
                  timestamp: nowStr,
                  user: activeUser,
                  action: 'Rollback Step',
                  details: `將關卡強制退回至 Step ${targetS.step} [${targetS.name}]。原因：${reason}`
                });
                sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
              }

              return { success: true, message: `已成功將專案退回至 Step ${rollbackStepNum}` };
            }
          }
        }
      }
    }
    throw new Error('找不到對應專案或子項目');
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 Phase C: 重啟已結案專案 (Restart Project)
// ==========================================
function api_restartProject(jobNumber, deliverableId, resetToStep1, reason) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('textjobtype') || h.includes('audit'));

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (currentJob.toLowerCase() === String(jobNumber).toLowerCase().trim() || !jobNumber) {

        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.includes('deliverables') && cellStr.includes(deliverableId)) {
            let wfData = JSON.parse(cellStr);
            let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);

            if (targetD && targetD.workflow) {
              const nowIso = new Date().toISOString();

              targetD.status = 'In Progress';
              if (idxStatus >= 0) sheet.getRange(i + 1, idxStatus + 1).setValue('In Progress');

              if (resetToStep1) {
                // 完全重置回 Step 1
                targetD.currentStep = 1;
                targetD.workflow.forEach((s, idx) => {
                  if (idx === 0) {
                    s.status = 'In Progress';
                    s.startedAt = nowIso;
                    s.completedAt = null;
                    s.completedAtIso = null;
                    s.remarks = `【專案重啟】${reason}`;
                  } else {
                    s.status = 'Pending';
                    s.startedAt = null;
                    s.pendingAssignmentAt = null;
                    s.completedAt = null;
                    s.completedAtIso = null;
                    s.reviewStatus = null;
                    s.remarks = '';
                  }
                });
              } else {
                // 保留現有進度，僅將最後一關重開啟
                let lastStep = targetD.workflow[targetD.workflow.length - 1];
                if (lastStep) {
                  lastStep.status = 'In Progress';
                  lastStep.startedAt = nowIso;
                  lastStep.completedAt = null;
                  lastStep.completedAtIso = null;
                  lastStep.remarks = `【專案重啟】${reason}`;
                  targetD.currentStep = lastStep.step;
                }
              }

              sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));

              if (idxAudit >= 0) {
                let logs = [];
                let cellStrLog = String(data[i][idxAudit] || '').trim();
                if (cellStrLog.startsWith('[')) {
                  try { logs = JSON.parse(cellStrLog); } catch(e) {}
                }
                const activeUser = Session.getActiveUser().getEmail().split('@')[0];
                const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
                
                logs.unshift({
                  timestamp: nowStr,
                  user: activeUser,
                  action: 'Restart Project',
                  details: `重啟專案 [${targetD.name}] (${resetToStep1 ? '重置至 Step 1' : '恢復結案前狀態'})。原因：${reason}`
                });
                sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
              }

              return { success: true, message: '專案已成功重啟！' };
            }
          }
        }
      }
    }
    throw new Error('找不到對應專案或子項目');
  } catch (e) {
    return { success: false, message: e.message };
  }
}