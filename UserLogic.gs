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

// ==========================================
// 💡 [更新版] 儲存工作流範本 (支援完美重新命名)
// ==========================================
function api_saveTemplate(newName, steps, originalName) {
  try {
    let sheet = getTemplateSheet();
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('WorkflowTemplates');
      sheet.appendRow(['TemplateName', 'WorkflowJSON']);
    }

    const data = sheet.getDataRange().getValues();
    const jsonStr = JSON.stringify(steps);
    let rowIndex = -1;

    // 1. 如果有提供舊名字 (代表是編輯模式)，先找舊名字的行數
    if (originalName) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(originalName).trim()) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    // 2. 如果沒找到，或沒有舊名字 (代表是新建模式)，就用新名字找
    if (rowIndex === -1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(newName).trim()) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    // 3. 覆蓋或新增
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1).setValue(newName); // 更新名稱
      sheet.getRange(rowIndex, 2).setValue(jsonStr); // 更新流程 JSON
    } else {
      sheet.appendRow([newName, jsonStr]);
    }
    return { success: true, message: '範本儲存成功！' };
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
// 💡 [純淨版] 建立新專案 (處理 TBD)
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
      delivList = [{ name: payload.deliverableName || '任務 1', templateName: templateName, templateJson: workflowSteps.length > 0 ? JSON.stringify(workflowSteps) : '' }];
    }

    const launchDeadlineStr = payload.deadline || payload.launchDate || Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
    const mainStartDateStr = payload.tentativeDate || Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

    let deliverables = delivList.map((d, idx) => {
      let steps = [];
      let tplName = d.templateName || d.type || payload.productType || '標準 Advertorial';

      // 💡 處理 TBD 狀態
      let delivDeadline = d.deadline || launchDeadlineStr;
      let delivStartDate = d.tentativeStartDate || mainStartDateStr;
      let isTBD = !!d.isTBD;
      
      if (d.deadline === '稍後補充' || isTBD) delivDeadline = '稍後補充';
      if (d.tentativeStartDate === '稍後補充' || isTBD) delivStartDate = '稍後補充';

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
          status: 'Pending Start',
          isStarted: false,
          fields: s.fields || ['URL'],
          parallelGroup: s.parallelGroup || s.group || '',
          keyDate: delivDeadline,    // 寫入子項目獨立死線
          deadline: delivDeadline,   // 寫入子項目獨立死線
          accumulatedDays: 0
        };
      });

      return {
        id: 'deliv_' + (idx + 1) + '_' + new Date().getTime(),
        name: d.name || ('篇章/任務 ' + (idx + 1)),
        type: tplName,
        status: 'Pending Start',
        currentStep: 1,
        tentativeStartDate: delivStartDate,
        mainDeadline: delivDeadline,
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
    if (idxStatus >= 0) rowData[idxStatus] = 'Pending Start'; 
    if (idxProduct >= 0) rowData[idxProduct] = JSON.stringify(workflowData);
    if (idxLog >= 0) rowData[idxLog] = JSON.stringify(auditLog);

    sheet.appendRow(rowData);

    if (typeof notifyFirebaseUpdate === 'function') {
      notifyFirebaseUpdate(payload.jobNumber, Session.getActiveUser().getEmail(), masterPmName, true);
    }

    return { success: true, message: '專案建立成功！' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [智能同步版] 修改單一關卡死線 (支援批量同步 TBD 關卡)
// ==========================================
function api_updateStepDeadline(deliverableId, stepNumber, newDeadline, cascade) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const activeUser = userEmail ? userEmail.split('@')[0] : 'System';

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('audit'));
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');

    const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");

    for (let i = 1; i < data.length; i++) {
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables') && cellStr.includes(deliverableId)) {
          let wfData = JSON.parse(cellStr);
          let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);

          if (targetD && targetD.workflow) {
            let logDetails = '';

            // 💡 智能批量同步：如果是 TBD 且前端按下確定，就把所有「未完成」且為「稍後補充」的關卡全部更新
            if (cascade) {
              targetD.workflow.forEach(s => {
                if (s.status !== 'Completed' && (s.keyDate === '稍後補充' || s.deadline === '稍後補充')) {
                  s.keyDate = newDeadline;
                  s.deadline = newDeadline;
                }
              });
              targetD.mainDeadline = newDeadline;
              logDetails = `將子項目 [${targetD.name}] 內所有「稍後補充 (TBD)」的關卡死線，批量同步為 [${newDeadline}]`;
            } else {
              // 僅更新單一關卡
              let targetS = targetD.workflow.find(s => s.step === parseFloat(stepNumber));
              if (targetS) {
                let oldDeadline = targetS.keyDate || targetS.deadline || '未設定';
                targetS.keyDate = newDeadline;
                targetS.deadline = newDeadline;
                logDetails = `將 Step ${targetS.step} [${targetS.name}] 的關卡死線由 [${oldDeadline}] 調整為 [${newDeadline}]`;
              }
            }

            sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));

            // 寫入 Audit Log
            if (idxAudit >= 0) {
              let logs = [];
              let cellStrLog = String(data[i][idxAudit] || '').trim();
              if (cellStrLog.startsWith('[')) { try { logs = JSON.parse(cellStrLog); } catch(e) {} }

              logs.unshift({
                timestamp: nowStr,
                user: activeUser,
                action: 'Update Step Deadline',
                details: logDetails
              });
              sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
            }

            let jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '') : '';
            if (typeof notifyFirebaseUpdate === 'function') {
              notifyFirebaseUpdate(jobNumber, userEmail, activeUser, true);
            }

            return { success: true, message: '關卡死線更新成功！' };
          }
        }
      }
    }
    throw new Error('找不到對應關卡');
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [UX 優化版] 我的任務看板數據 (保證啟動後留在「進行中」區域)
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
      try {
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
          const deliverableCount = wfData.deliverables.length;

          wfData.deliverables.forEach((d, dIdx) => {
            if (d.status === 'Completed' || d.status === 'Deleted' || d.status === 'Recycle Bin') return;

            const displayJobNum = (deliverableCount > 1 && rowJobNum) ? `${rowJobNum}-P${dIdx + 1}` : rowJobNum;

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

              if (userRole === 'Team Head') {
                activeMatch = activeSteps.some(s => s.dept === userDept);
                futureMatch = d.workflow && d.workflow.some(s => s.step > d.currentStep && s.dept === userDept);
              } else {
                activeMatch = activeSteps.some(s => s.assignee && s.assignee.toLowerCase() === userName.toLowerCase());
                futureMatch = d.workflow && d.workflow.some(s => s.step > d.currentStep && s.assignee && s.assignee.toLowerCase() === userName.toLowerCase());
              }

              if (activeMatch) {
                isMyTask = true;
              } else if (futureMatch) {
                isMyTask = true;
                isUpcomingForUser = true; 
              }
            }

            if (!isMyTask) return; 

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
            let daysLeft = 999;
            if (!isNaN(deadlineDate.getTime())) {
              daysLeft = Math.ceil((deadlineDate - today) / (1000 * 60 * 60 * 24));
            }

            const item = {
              jobNumber: displayJobNum, 
              client: clientVal, taskName: d.name || '未命名任務',
              type: d.type || 'Standard', status: d.status || 'Pending Start',
              deadline: effectiveDeadlineStr, daysLeft: daysLeft,
              currentStepName: currentStepName, assignee: currentAssignee,
              pmName: pmName, salesName: salesName
            };

            let isPaused = d.status === 'Paused' || pStatus.includes('pause');
            
            // 💡 UX 優化：只要專案已經啟動，即使下一關還在等待派案，也算在進行中，避免退回待啟動區
            let isProjectStarted = d.status !== 'Pending Start' && d.status !== 'Not Started';
            let isActive = primaryActiveStep && (primaryActiveStep.status === 'In Progress' || (isProjectStarted && primaryActiveStep.status === 'Pending'));

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
      } catch (rowErr) {}
    }
    
    const sortByDeadline = (a, b) => {
      let dA = new Date(a.deadline).getTime();
      let dB = new Date(b.deadline).getTime();
      if (isNaN(dA)) dA = 9999999999999;
      if (isNaN(dB)) dB = 9999999999999;
      return dA - dB;
    };
    result.overdue.sort(sortByDeadline);
    result.dueSoon.sort(sortByDeadline);
    result.onTrack.sort(sortByDeadline);
    result.upcoming.sort(sortByDeadline);
    result.paused.sort(sortByDeadline);

    return JSON.parse(JSON.stringify({ success: true, data: result }));
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 1. 取得單一專案工作流資料 (100% 自動對齊 pmName 欄位)
// ==========================================
function api_getProjectWorkflow(jobNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h === 'pmname' || h === 'pm');
    const idxSales = headers.findIndex(h => h.includes('salesperson') || h.includes('sales'));
    const idxDeadline = headers.findIndex(h => h.includes('launchdate') || h.includes('deadline'));
    const idxProd = headers.findIndex(h => h === 'productname' || h.includes('product'));
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('audit'));

   

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '';
        const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';
        const salesPerson = idxSales >= 0 ? String(data[i][idxSales] || '').trim() : '';
        const mainDeadline = cleanYMD(idxDeadline >= 0 ? data[i][idxDeadline] : '');

        let wfData = {};
        let cellStr = idxProd >= 0 ? String(data[i][idxProd] || '') : '';
        if (!cellStr.startsWith('{')) {
          for (let c = 0; c < data[i].length; c++) {
            let colVal = String(data[i][c] || '');
            if (colVal.includes('deliverables')) { cellStr = colVal; break; }
          }
        }
        if (cellStr.startsWith('{')) {
          try { wfData = JSON.parse(cellStr); } catch(e) {}
        }

        // 💡 [PM 讀取自動覆寫] 所有 dept === 'PM' 的關卡，100% 動態對齊 Projects 表內的 pmName
        if (pmName && wfData.deliverables) {
          wfData.deliverables.forEach(d => {
            if (d.workflow) {
              d.workflow.forEach(s => {
                if (String(s.dept || '').toUpperCase() === 'PM') {
                  s.assignee = pmName;
                }
              });
            }
          });
        }

        let auditLog = [];
        if (idxAudit >= 0) {
          let logStr = String(data[i][idxAudit] || '').trim();
          if (logStr.startsWith('[')) { try { auditLog = JSON.parse(logStr); } catch(e){} }
        }

        return {
          success: true,
          data: {
            jobNumber: jobNumber,
            client: clientName,
            pmName: pmName,
            salesPerson: salesPerson,
            mainDeadline: mainDeadline,
            workflowData: wfData,
            auditLog: auditLog
          }
        };
      }
    }
    throw new Error('找不到專案 ' + jobNumber);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 修正 api_dispatchWorkflowStep (全 PM 關卡雙向同步修復)
// ==========================================
function api_dispatchWorkflowStep(jobNumber, deliverableId, stepNumber, assignee) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const activeUser = userEmail ? userEmail.split('@')[0] : 'System';

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('audit'));
    const idxPM = headers.findIndex(h => h === 'pmname' || h === 'pm');

    const nowIso = new Date().toISOString();
    const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (currentJob.toLowerCase() === String(jobNumber).trim().toLowerCase()) {

        let cellColIdx = -1;
        let cellStr = '';
        for (let c = 0; c < data[i].length; c++) {
          let colVal = String(data[i][c] || '');
          if (colVal.includes('deliverables') && colVal.includes(deliverableId)) {
            cellStr = colVal;
            cellColIdx = c;
            break;
          }
        }

        if (cellColIdx >= 0 && cellStr.startsWith('{')) {
          let wfData = JSON.parse(cellStr);
          let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);

          if (targetD && targetD.workflow) {
            let targetS = targetD.workflow.find(s => s.step === parseFloat(stepNumber));
            if (targetS) {
              let oldAssignee = targetS.assignee || '未指派';
              let isPMDept = String(targetS.dept || '').toUpperCase() === 'PM';

              // (1) 無論在哪一個 PM 關卡點擊更換，一律同步更新 Projects 表內的 pmName 欄位
              if (isPMDept && idxPM >= 0) {
                sheet.getRange(i + 1, idxPM + 1).setValue(assignee);
              }

              // (2) 若更換的是 PM 關卡，將專案內所有屬於 PM 的關卡負責人同步更新為新 PM
              if (isPMDept) {
                (wfData.deliverables || []).forEach(d => {
                  (d.workflow || []).forEach(s => {
                    if (String(s.dept || '').toUpperCase() === 'PM') {
                      s.assignee = assignee;
                    }
                  });
                });
              } else {
                targetS.assignee = assignee;
              }

              let isCurrentTurn = (targetS.step === targetD.currentStep) || 
                                  targetD.workflow.filter(s => s.step < targetS.step && !s.parallelGroup).every(s => s.status === 'Completed');

              if (isCurrentTurn) {
                if (targetS.status === 'Pending Start' || targetS.status === 'Pending') {
                  targetS.status = 'In Progress';
                  targetS.startedAt = nowIso;
                  if (targetS.pendingAssignmentAt) {
                    targetS.dispatchWaitMs = (targetS.dispatchWaitMs || 0) + Math.max(0, new Date().getTime() - new Date(targetS.pendingAssignmentAt).getTime());
                    targetS.pendingAssignmentAt = null;
                  }
                }
              }

              sheet.getRange(i + 1, cellColIdx + 1).setValue(JSON.stringify(wfData));

              if (idxAudit >= 0) {
                let logs = [];
                let cellStrLog = String(data[i][idxAudit] || '').trim();
                if (cellStrLog.startsWith('[')) { try { logs = JSON.parse(cellStrLog); } catch(e) {} }

                let actionText = isPMDept 
                  ? `將專案 PM 及所有 PM 關卡負責人由 [${oldAssignee}] 更換為 [${assignee}]`
                  : (oldAssignee === '未指派'
                     ? `將 Step ${targetS.step} [${targetS.name}] 指派給 [${assignee}]`
                     : `將 Step ${targetS.step} 的負責人由 [${oldAssignee}] 更換為 [${assignee}]`);

                logs.unshift({ timestamp: nowStr, user: activeUser, action: 'Dispatch Task', details: actionText });
                sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
              }

              if (typeof notifyFirebaseUpdate === 'function') {
                notifyFirebaseUpdate(jobNumber, userEmail, activeUser, true);
              }

              return { success: true, message: '派案與 PM 資料同步成功！' };
            }
          }
        }
      }
    }
    throw new Error('找不到對應專案或關卡');
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 3. 更換 PM API (同步寫入 pmName 欄位並更新所有 PM 關卡)
// ==========================================
function api_reassignPM(jobNumber, newPM) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const activeUser = userEmail ? userEmail.split('@')[0] : 'System';

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxPM = headers.findIndex(h => h === 'pmname' || h === 'pm');
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('audit'));
    const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        
        let oldPM = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '未設定';

        // 💡 (1) 寫入 Projects 工作表欄位 pmName
        if (idxPM >= 0) {
          sheet.getRange(i + 1, idxPM + 1).setValue(newPM);
        }

        // 💡 (2) 找到 JSON，將所有 dept === 'PM' 的關卡同步為 newPM
        for (let c = 0; c < data[i].length; c++) {
          let cellStr = String(data[i][c] || '');
          if (cellStr.startsWith('{') && cellStr.includes('deliverables')) {
            let wfData = JSON.parse(cellStr);
            if (wfData.deliverables) {
              wfData.deliverables.forEach(d => {
                if (d.workflow) {
                  d.workflow.forEach(s => {
                    if (String(s.dept || '').toUpperCase() === 'PM') {
                      s.assignee = newPM;
                    }
                  });
                }
              });
              sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));
            }
            break;
          }
        }

        if (idxAudit >= 0) {
          let logs = [];
          let logStr = String(data[i][idxAudit] || '').trim();
          if (logStr.startsWith('[')) { try { logs = JSON.parse(logStr); } catch(e){} }

          logs.unshift({
            timestamp: nowStr,
            user: activeUser,
            action: 'Reassign PM',
            details: `將專案負責 PM 由 [${oldPM}] 更換為 [${newPM}]，已自動同步所有 PM 關卡`
          });
          sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
        }

        if (typeof notifyFirebaseUpdate === 'function') {
          notifyFirebaseUpdate(jobNumber, userEmail, activeUser, true);
        }

        return { success: true, message: 'PM 更換成功，所有 PM 關卡已自動同步！' };
      }
    }
    throw new Error('找不到專案 ' + jobNumber);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 取得已完成專案歷史紀錄 (修復版：精準讀取 JSON 已結案子項目與 10 大欄位)
// ==========================================
function api_getCompletedProjects(keyword, startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum   = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient   = headers.findIndex(h => h.includes('client'));
    const idxPM       = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxDeadline = headers.findIndex(h => h.includes('launchdate') || h.includes('deadline'));
    const idxProd     = headers.findIndex(h => h === 'productname' || h.includes('product'));
    const idxStatus   = headers.findIndex(h => h === 'status' || h === 'project_status');

    let list = [];
    const kw = String(keyword || '').trim().toLowerCase();

 

    for (let i = 1; i < data.length; i++) {
      const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '';
      const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';
      const mainDeadline = cleanYMD(idxDeadline >= 0 ? data[i][idxDeadline] : '');
      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';

      let cellStr = idxProd >= 0 ? String(data[i][idxProd] || '') : '';
      if (!cellStr.startsWith('{')) {
        for (let c = 0; c < data[i].length; c++) {
          let colVal = String(data[i][c] || '');
          if (colVal.includes('deliverables')) { cellStr = colVal; break; }
        }
      }

      let wfData = {};
      if (cellStr.startsWith('{')) {
        try { wfData = JSON.parse(cellStr); } catch(e) {}
      }

      if (wfData && wfData.deliverables && wfData.deliverables.length > 0) {
        wfData.deliverables.forEach(d => {
          // 💡 精準判定：子項目狀態為 Completed，或專案總狀態為 Completed
          if (d.status === 'Completed' || pStatus === 'Completed') {
            let revCount = d.revisionCount || 0;
            if (!revCount && d.workflow) {
              revCount = d.workflow.filter(s => s.isSubStep || (s.name && s.name.includes('退回修改'))).length;
            }

            // 關鍵字搜尋過濾
            if (kw) {
              let match = jobNumber.toLowerCase().includes(kw) || 
                          clientName.toLowerCase().includes(kw) || 
                          (d.name || '').toLowerCase().includes(kw) || 
                          pmName.toLowerCase().includes(kw);
              if (!match) return;
            }

            let compAtStr = d.completedAt || '稍早';
            let launchUrl = '';
            if (d.workflow && d.workflow.length > 0) {
              let lastStep = d.workflow[d.workflow.length - 1];
              if (lastStep.submittedData) {
                launchUrl = lastStep.submittedData['正式上線網址 / 連結'] || lastStep.submittedData['URL'] || '';
              }
            }

            let totalMs = 0;
            (d.workflow || []).forEach(s => {
              if (s.dispatchWaitMs) totalMs += s.dispatchWaitMs;
              if (s.accumulatedMs) totalMs += s.accumulatedMs;
            });

            let daysDiff = 0;
            if (mainDeadline && mainDeadline !== '未設定' && compAtStr.length >= 10) {
              let d1 = new Date(mainDeadline);
              let d2 = new Date(compAtStr.split(' ')[0]);
              daysDiff = Math.ceil((d2 - d1) / (1000 * 60 * 60 * 24));
            }

            let compDateOnly = compAtStr.split(' ')[0];
            if (startDate && compDateOnly < startDate) return;
            if (endDate && compDateOnly > endDate) return;

            list.push({
              jobNumber: jobNumber,
              client: clientName,
              deliverableName: d.name || '未命名任務',
              pmName: pmName,
              totalDurationText: formatDurationText(totalMs),
              deadline: mainDeadline,
              completedAt: compAtStr,
              revisionCount: revCount,
              launchUrl: launchUrl,
              isDelayed: daysDiff > 0,
              delayText: daysDiff > 0 ? `延誤 ${daysDiff} 天` : '按時完成'
            });
          }
        });
      }
    }

    return { success: true, data: list };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

    

// ==========================================
// 💡 更新工作流狀態 API (修正版：完美身分同步，斷絕讀取時間差)
// ==========================================
function api_updateWorkflowState(jobNumber, deliverableId, payload) {
  try {
    const userEmail = payload.userEmail || Session.getActiveUser().getEmail();
    const activeUser = userEmail ? userEmail.split('@')[0] : 'System';

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxProd   = headers.findIndex(h => h === 'productname' || h.includes('product'));
    const idxAudit  = headers.findIndex(h => h === 'textjobtype' || h.includes('audit'));

    const now = new Date();
    const nowIso = now.toISOString();
    const nowStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm");

    for (let i = 1; i < data.length; i++) {
      const currentJob = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (currentJob.toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        
        let cellColIdx = -1;
        let cellStr = idxProd >= 0 ? String(data[i][idxProd] || '') : '';
        if (cellStr.startsWith('{')) {
          cellColIdx = idxProd;
        } else {
          for (let c = 0; c < data[i].length; c++) {
            let colVal = String(data[i][c] || '');
            if (colVal.includes('deliverables') && colVal.includes(deliverableId)) {
              cellStr = colVal;
              cellColIdx = c;
              break;
            }
          }
        }

        if (cellColIdx < 0 || !cellStr.startsWith('{')) {
          throw new Error('找不到對應子項目的工作流 JSON 資料');
        }

        let wfData = JSON.parse(cellStr);
        let deliverable = (wfData.deliverables || []).find(d => d.id === deliverableId);
        if (!deliverable) throw new Error('找不到對應子項目');

        const action = payload.action;
        const stepNum = parseFloat(payload.stepNumber);

        if (action === 'START') {
          deliverable.status = 'In Progress';
          deliverable.startedAt = nowIso;
          if (deliverable.workflow && deliverable.workflow.length > 0) {
            let firstS = deliverable.workflow[0];
            firstS.status = firstS.assignee ? 'In Progress' : 'Pending';
            firstS.startedAt = firstS.assignee ? nowIso : null;
            firstS.pendingAssignmentAt = firstS.assignee ? null : nowIso;
            firstS.isStarted = true;
          }
        } else if (action === 'SUBMIT') {
          let currentStepObj = deliverable.workflow.find(s => s.step === stepNum);
          if (currentStepObj) {
            currentStepObj.status = 'Completed';
            currentStepObj.completedAt = nowStr;
            currentStepObj.completedAtIso = nowIso;
            if (payload.inputs) {
              currentStepObj.submittedData = payload.inputs;
            }

            if (currentStepObj.startedAt) {
              let startMs = new Date(currentStepObj.startedAt).getTime();
              currentStepObj.accumulatedMs = (currentStepObj.accumulatedMs || 0) + Math.max(0, now.getTime() - startMs);
            }

            let nextStepObj = deliverable.workflow.find(s => s.step > stepNum && s.status !== 'Completed');
            if (nextStepObj) {
              deliverable.currentStep = nextStepObj.step;
              if (nextStepObj.assignee) {
                nextStepObj.status = 'In Progress';
                nextStepObj.startedAt = nowIso;
                nextStepObj.pendingAssignmentAt = null;
              } else {
                // 💡 關鍵修復：將狀態改為 Pending (等待中)，而不是舊的 Pending Start
                nextStepObj.status = 'Pending'; 
                nextStepObj.pendingAssignmentAt = nowIso;
              }
            } else {
              deliverable.status = 'Completed';
              deliverable.completedAt = nowStr;
              deliverable.completedAtIso = nowIso;
            }
          }
        }

        sheet.getRange(i + 1, cellColIdx + 1).setValue(JSON.stringify(wfData));

        let allDeliverablesCompleted = (wfData.deliverables || [])
          .filter(d => d.status !== 'Deleted' && d.status !== 'Recycle Bin')
          .every(d => d.status === 'Completed');

        if (allDeliverablesCompleted && idxStatus >= 0) {
          sheet.getRange(i + 1, idxStatus + 1).setValue('Completed'); 
        } else if (idxStatus >= 0 && String(data[i][idxStatus]).trim() === 'Pending Start') {
          sheet.getRange(i + 1, idxStatus + 1).setValue('In Progress');
        }

        if (idxAudit >= 0) {
          let logs = [];
          let logStr = String(data[i][idxAudit] || '').trim();
          if (logStr.startsWith('[')) { try { logs = JSON.parse(logStr); } catch(e){} }

          let logDetail = action === 'START' 
            ? `啟動了子項目 [${deliverable.name}]` 
            : `完成了 Step ${stepNum} 的關卡工作` + (deliverable.status === 'Completed' ? ' (全案完工)' : '');

          logs.unshift({ timestamp: nowStr, user: activeUser, action: action, details: logDetail });
          sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
        }

        if (typeof notifyFirebaseUpdate === 'function') {
          notifyFirebaseUpdate(jobNumber, userEmail, activeUser, true);
        }

        // 💡 關鍵修復：直接將更新好的 wfData 回傳
        return { success: true, message: '工作流狀態更新成功！', updatedWorkflowData: wfData };
      }
    }
    throw new Error('找不到專案 ' + jobNumber);
  } catch (e) {
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

// ==========================================
// 💡 取得專案回收箱列表 (支援專案層級 Column E 與子項目層級雙軌撈取)
// ==========================================
function api_getRecycleBinProjects() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM     = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxProd   = headers.findIndex(h => h === 'productname' || h.includes('product'));

    let recycleList = [];

    for (let i = 1; i < data.length; i++) {
      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
      const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : data[i][0];
      const client = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '';
      const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';

      let isProjectRecycled = (pStatus.toLowerCase() === 'recycle bin' || pStatus.toLowerCase() === 'deleted');

      if (isProjectRecycled) {
        recycleList.push({
          jobNumber: jobNumber,
          client: client,
          pmName: pmName,
          taskName: '整項專案'
        });
      } else {
        // 檢查子項目層級是否有被單獨刪除/移至回收箱
        let cellStr = idxProd >= 0 ? String(data[i][idxProd] || '') : '';
        if (cellStr.startsWith('{')) {
          try {
            let wfData = JSON.parse(cellStr);
            if (wfData && wfData.deliverables) {
              wfData.deliverables.forEach(d => {
                if (d.status === 'Recycle Bin' || d.status === 'Deleted') {
                  recycleList.push({
                    jobNumber: jobNumber,
                    client: client,
                    pmName: pmName,
                    taskName: d.name || '子項目'
                  });
                }
              });
            }
          } catch(e){}
        }
      }
    }

    return { success: true, data: recycleList };
  } catch (e) {
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
// 💡 部門營運數據 API (去重優化版：1 項目 1 卡片 + 精準統計關卡數)
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
    
    // 💡 暫存檔：用於對 Completed 項目進行去重
    let completedMap = {};
    let totalCompletedStepsCount = 0; // 統計 5 天內完成的關卡總次數

    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

    for (let i = 1; i < data.length; i++) {
      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
      if (pStatus === 'Recycle Bin' || pStatus === 'Cancelled') continue;

      let wfData = {};
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables')) {
          try { wfData = JSON.parse(cellStr); break; } catch(e){}
        }
      }

      if (wfData && wfData.deliverables) {
        const deliverableCount = wfData.deliverables.length; 

        wfData.deliverables.forEach((d, dIdx) => { 
          if (d.status === 'Deleted') return;

          const baseJobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '') : '';
          const jobNumber = (deliverableCount > 1 && baseJobNumber) ? `${baseJobNumber}-P${dIdx + 1}` : baseJobNumber; 
          
          const client = idxClient >= 0 ? String(data[i][idxClient] || '') : '';
          const pmName = idxPM >= 0 ? String(data[i][idxPM] || '') : '';

          if (!d.workflow || d.workflow.length === 0) return;

          // ----------------------------------------------------
          // 💡 1. 採計產能與已完成關卡（進行去重，1 項目僅留 1 卡片）
          // ----------------------------------------------------
          d.workflow.forEach(s => {
            const sDept = String(s.dept || '').trim().toLowerCase();
            const assignee = String(s.assignee || '').trim();

            if (sDept === targetDept.toLowerCase()) {
              if (assignee) {
                if (!capacityMap[assignee]) {
                  capacityMap[assignee] = { name: assignee, inProgress: 0, completed: 0, revisions: 0 };
                }
                if (s.status === 'In Progress') capacityMap[assignee].inProgress++;
                if (s.status === 'Completed') capacityMap[assignee].completed++;
              }

              if (s.status === 'Completed') {
                let compDateStr = todayStr;
                if (s.completedAt) compDateStr = String(s.completedAt).split(' ')[0];
                else if (s.completedAtIso) compDateStr = String(s.completedAtIso).split('T')[0];

                let compDate = new Date(compDateStr);
                let diffDays = Math.floor((now - compDate) / (1000 * 60 * 60 * 24));

                if (isNaN(diffDays) || diffDays <= 5) {
                  totalCompletedStepsCount++; // 關卡總次數 +1

                  const delivKey = jobNumber + '_' + (d.id || dIdx);
                  const compTask = {
                    jobNumber: jobNumber, 
                    client: client,
                    taskName: d.name || '未命名任務',
                    stepNumber: s.step,
                    stepName: s.name || ('Step ' + s.step),
                    assignee: assignee || pmName,
                    deliverableId: d.id,
                    deadline: compDateStr,
                    category: 'COMPLETED'
                  };

                  // 💡 去重覆蓋：同個項目若有多個關卡完成，只保留最新完成的那一個關卡資訊
                  completedMap[delivKey] = compTask;
                }
              }
            }
          });

          // ----------------------------------------------------
          // 💡 2. 處理進行中 (In Progress) 與預備管線 (Pipeline)
          // ----------------------------------------------------
          let activeStep = d.workflow.find(s => s.status === 'In Progress');
          const allCompleted = d.status === 'Completed' || d.workflow.every(s => s.status === 'Completed');

          if (!activeStep && !allCompleted && d.workflow.length > 0) {
            activeStep = d.workflow.find(s => s.step === d.currentStep) || d.workflow[0];
          }

          if (activeStep && !allCompleted) {
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
              else if (!assignee || assignee === '未指派') taskCategory = 'UNASSIGNED';
              else {
                let daysLeft = Math.ceil((new Date(deadline) - new Date(todayStr)) / 86400000);
                if (daysLeft <= 2) taskCategory = 'DUE_SOON';
                else if (activeStep.status === 'Pending Start' || activeStep.status === 'Pending') taskCategory = 'PIPELINE';
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

              if (deadline && taskCategory !== 'PIPELINE') {
                calendarTasks.push(taskObj);
                calendarMap[deadline] = (calendarMap[deadline] || 0) + 1;
              }

              if (taskCategory === 'CLIENT_REVIEW') clientReviewTasks.push(taskObj);
              else if (taskCategory === 'OVERDUE' || taskCategory === 'UNASSIGNED') riskTasks.push(taskObj);
              else if (taskCategory === 'PIPELINE') {
                pipelineTasks.push({
                  jobNumber: jobNumber, 
                  client: client,
                  taskName: d.name || '未命名任務',
                  stepNumber: activeStep.step,
                  stepName: currentRealStepName,
                  assignee: assignee || pmName,
                  frontDept: 'PM',
                  currentFrontStep: '待啟動 (Pending Start)',
                  frontAssignee: pmName || '未指派',
                  estimatedArrival: deadline,
                  pmName: pmName,
                  deliverableId: d.id
                });
              }
              else activeTasks.push(taskObj);

            } else {
              const futureStep = d.workflow.find(s => s.step > activeStep.step && String(s.dept || '').trim().toLowerCase() === targetDept.toLowerCase() && s.status !== 'Completed');
              if (futureStep) {
                pipelineTasks.push({
                  jobNumber: jobNumber, 
                  client: client,
                  taskName: d.name || '未命名任務',
                  stepNumber: futureStep.step,
                  stepName: futureStep.name,
                  assignee: futureStep.assignee || '未指定',
                  frontDept: activeStep.dept || 'PM',
                  currentFrontStep: currentRealStepName + ' (' + activeStep.dept + ')',
                  frontAssignee: activeStep.assignee || (activeStep.dept + ' 未指派'),
                  estimatedArrival: deadline,
                  pmName: pmName,
                  deliverableId: d.id
                });
              }
            }
          }
        });
      }
    }

    // 將去重後的物件轉為陣列
    const completedTasks = Object.keys(completedMap).map(k => completedMap[k]);
    
    // 把去重後的已完成卡片併入日曆呈現
    completedTasks.forEach(ct => {
      calendarTasks.push(ct);
      calendarMap[ct.deadline] = (calendarMap[ct.deadline] || 0) + 1;
    });

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
        completedTasks: completedTasks,
        totalCompletedStepsCount: totalCompletedStepsCount // 💡 回傳真實關卡完成次數
      }
    };
  } catch (e) {
    return { success: false, message: 'api_getDeptOperationData 錯誤: ' + e.message };
  }
}



// ==========================================
// 💡 補齊 API：修改專案總死線 (Launch Date) + 智能同步 JSON 子項目
// ==========================================
function api_updateProjectDeadline(jobNumber, newDeadline) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) throw new Error('找不到 Projects 工作表');

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxLaunch = headers.findIndex(h => h.includes('launch') || h.includes('deadline') || h.includes('死線'));
    const idxLog = headers.findIndex(h => h.includes('textjobtype') || h.includes('log') || h.includes('audit'));
    const idxProd = headers.findIndex(h => h === 'productname' || h.includes('product') || h.includes('workflow'));

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

    // 1. 更新試算表主欄位
    sheet.getRange(rowIndex, idxLaunch + 1).setValue(newDeadline);

    // 2. 💡 關鍵修復：同步更新 JSON 內的子項目死線
    if (idxProd >= 0) {
      let cellStr = String(data[rowIndex - 1][idxProd] || '');
      if (cellStr.startsWith('{')) {
        try {
          let wfData = JSON.parse(cellStr);
          if (wfData && wfData.deliverables) {
            wfData.deliverables.forEach(d => {
              // 如果子項目的死線本來是「稍後補充」，就跟著主死線一起更新！
              if (d.mainDeadline === '稍後補充' || !d.mainDeadline) {
                d.mainDeadline = newDeadline;
                if (d.workflow) {
                  d.workflow.forEach(s => {
                    // 同步更新底下尚未完成的 TBD 關卡
                    if (s.status !== 'Completed' && (s.keyDate === '稍後補充' || s.deadline === '稍後補充')) {
                      s.keyDate = newDeadline;
                      s.deadline = newDeadline;
                    }
                  });
                }
              }
            });
            // 存回 JSON
            sheet.getRange(rowIndex, idxProd + 1).setValue(JSON.stringify(wfData));
          }
        } catch(e) { 
          console.error('JSON 同步解析失敗:', e); 
        }
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
        action: 'Update Deadline',
        details: `將專案總死線修改為：${newDeadline} (已同步更新 TBD 子項目)`
      });

      sheet.getRange(rowIndex, idxLog + 1).setValue(JSON.stringify(logArray));
    }

    // 4. 發送更新廣播
    if (typeof notifyFirebaseUpdate === 'function') {
      const uEmail = Session.getActiveUser().getEmail();
      notifyFirebaseUpdate(jobNumber, uEmail, uEmail.split('@')[0], true);
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
    // 💡 翻譯蒟蒻在這裡！把 1A-P1 的尾巴砍掉，變回 1A
    jobNumber = String(jobNumber).split('-P')[0];

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

// ==========================================
// 💡 Phase D: 專屬 Client 中繼節點的「無限修改展開引擎」
// ==========================================
function api_triggerDynamicClientRevision(jobNumber, deliverableId, stepNumber, feedback) {
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
              
              let targetIdx = targetD.workflow.findIndex(s => s.step === parseFloat(stepNumber));
              let targetS = targetD.workflow[targetIdx];
              if (!targetS) throw new Error('找不到目標退回關卡');

              // 尋找最近一個非 Client 的「製作關卡」以繼承部門名稱
              let prevS = targetD.workflow.slice(0, targetIdx).reverse().find(s => s.dept !== 'Client');
              if (!prevS) prevS = { dept: 'Editorial', name: '前置作業' };

              const nowIso = new Date().toISOString();
              const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");

              // 1. 將當前的 Client 關卡標記為完成 (因為審批動作本身結束了，進入修改)
              targetS.status = 'Completed';
              targetS.completedAt = nowStr;
              targetS.completedAtIso = nowIso;
              targetS.submittedData = { "客戶意見": feedback };
              
              if (targetS.startedAt) {
                targetS.accumulatedMs = (targetS.accumulatedMs || 0) + Math.max(0, new Date().getTime() - new Date(targetS.startedAt).getTime());
              }

              // 2. 小數點推算法：生成 3a 與 3b (或 3.01, 3.02)
              let baseStep = Math.floor(targetS.step); 
              let subStepCount = targetD.workflow.filter(s => Math.floor(s.step) === baseStep).length;
              
              let newStepA_num = baseStep + (subStepCount * 0.01);
              let newStepB_num = baseStep + ((subStepCount + 1) * 0.01);

              let stepA_Editor = {
                step: newStepA_num,
                name: `[退回修改] ${prevS.name}`,
                dept: prevS.dept,
                assignee: targetS.assignee, // 由原負責聯絡的同事執行修改
                status: 'In Progress',
                startedAt: nowIso,
                dispatchWaitMs: 0,
                remarks: `【客戶修改要求】${feedback}`,
                isSubStep: true,
                baseStep: baseStep,
                parallelGroup: targetS.parallelGroup || ''
              };

              let stepB_Client = {
                step: newStepB_num,
                name: `Client 再次審批`,
                dept: 'Client',
                assignee: targetS.assignee,
                status: 'Pending',
                startedAt: null,
                isSubStep: true,
                baseStep: baseStep,
                parallelGroup: targetS.parallelGroup || ''
              };

              // 3. 動態陣列插隊 (splice)
              targetD.workflow.splice(targetIdx + 1, 0, stepA_Editor, stepB_Client);
              
              // 確保順序正確
              targetD.workflow.sort((a, b) => a.step - b.step);
              targetD.currentStep = stepA_Editor.step;

              // 4. 退件計數器 (奧客警報數據源🚨)
              targetD.revisionCount = (targetD.revisionCount || 0) + 1;

              sheet.getRange(i + 1, c + 1).setValue(JSON.stringify(wfData));

              // 5. 寫入日誌
              if (idxAudit >= 0) {
                let logs = [];
                let cellStrLog = String(data[i][idxAudit] || '').trim();
                if (cellStrLog.startsWith('[')) { try { logs = JSON.parse(cellStrLog); } catch(e) {} }
                const activeUser = Session.getActiveUser().getEmail().split('@')[0];
                
                let alertStr = targetD.revisionCount >= 5 ? ' 🚨[奧客警報: 第 '+targetD.revisionCount+' 次退件]' : '';

                logs.unshift({
                  timestamp: nowStr,
                  user: activeUser,
                  action: `Client Revision${alertStr}`,
                  details: `客戶退回修改，新增關卡 [修改 ${prevS.name}]。意見：${feedback}`
                });
                sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));
              }

              return { success: true, message: `已產生修改流程，此項目為第 ${targetD.revisionCount} 次退回！` };
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

// 💡 輔助函數：將毫秒轉換為 天/小時/分鐘 (用於後端)
function formatMsToText(ms) {
  if (!ms || ms <= 0) return '少於 1 分鐘';
  let totalMins = Math.floor(ms / (1000 * 60));
  let mins = totalMins % 60;
  let totalHours = Math.floor(totalMins / 60);
  let hours = totalHours % 24;
  let days = Math.floor(totalHours / 24);
  let res = '';
  if (days > 0) res += days + '天 ';
  if (hours > 0) res += hours + '小時 ';
  res += mins + '分鐘';
  return res.trim();
}

function api_sendNudge(jobNumber, deliverableId, stepNumber) {
  try {
    // 💡 翻譯蒟蒻在這裡！把 1A-P1 的尾巴砍掉，變回 1A
    jobNumber = String(jobNumber).split('-P')[0]; 
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('textjobtype') || h.includes('audit'));
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxPM = headers.findIndex(h => h === 'pm' || h.includes('pmname')); 

    for (let i = 1; i < data.length; i++) {
      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
      if (pStatus === 'Recycle Bin' || pStatus === 'Deleted') continue; 

      if (String(data[i][idxJobNum >= 0 ? idxJobNum : 0]).trim().toLowerCase() === String(jobNumber).trim().toLowerCase()) {
        
        if (idxAudit >= 0) {
          let logs = [];
          let cellStrLog = String(data[i][idxAudit] || '').trim();
          if (cellStrLog.startsWith('[')) { try { logs = JSON.parse(cellStrLog); } catch(e) {} }
          
          const activeUser = Session.getActiveUser().getEmail().split('@')[0];
          const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
          
          logs.unshift({ timestamp: nowStr, user: activeUser, action: 'Nudge', details: `📢 針對 Step ${stepNumber} 發出了緊急催辦提醒！` });
          sheet.getRange(i + 1, idxAudit + 1).setValue(JSON.stringify(logs));

          let sentTo = "未發送"; 

          for (let c = 0; c < data[i].length; c++) {
            let cellStrDeliv = String(data[i][c] || '');
            if (cellStrDeliv.includes('deliverables') && cellStrDeliv.includes(deliverableId)) {
              try {
                let wfData = JSON.parse(cellStrDeliv);
                let targetD = (wfData.deliverables || []).find(d => d.id === deliverableId);
                
                if (targetD && targetD.workflow) {
                  let targetS = targetD.workflow.find(s => s.step === parseFloat(stepNumber));
                  if (targetS) {
                    
                    let realAssignee = targetS.assignee;
                    if (!realAssignee && String(targetS.dept).toUpperCase() === 'PM' && idxPM >= 0) {
                      realAssignee = String(data[i][idxPM] || '').trim();
                    }

                    if (realAssignee && typeof sendSystemEmail === 'function') {
                      sentTo = sendSystemEmail(
                        realAssignee, 
                        `[緊急催辦] 專案 ${jobNumber} 任務提醒`, 
                        `🚨 緊急催辦通知`, 
                        `專案負責人剛剛針對 <b>[${jobNumber}]</b> 的 <b>[${targetS.name}]</b> 關卡發出了緊急催辦，請立即跟進！`
                      );
                    } else {
                      sentTo = "無負責人";
                    }
                  } 
                } 
              } catch(e) {}
              break; 
            }
          }
          return { success: true, message: '催辦已寫入！(信件發送至: ' + sentTo + ')' };
        }
      }
    }
    throw new Error('找不到對應專案');
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 💡 [新增] 混合式動態通知引擎 (Notification Center)
// ==========================================
function api_getUserNotifications(userEmail) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) return { success: true, data: [] };

    const activeEmail = userEmail || Session.getActiveUser().getEmail();
    const userName = activeEmail.split('@')[0].toLowerCase();
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('textjobtype') || h.includes('audit'));

    let notifications = [];
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
    const sevenDaysAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000)); // 只抓最近 7 天的日誌

    for (let i = 1; i < data.length; i++) {
      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
      if (pStatus === 'Recycle Bin' || pStatus === 'Deleted') continue;

      const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '';
      const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim().toLowerCase() : '';

      let wfData = {};
      let logs = [];
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables')) { try { wfData = JSON.parse(cellStr); } catch(e){} }
        if (c === idxAudit && cellStr.startsWith('[')) { try { logs = JSON.parse(cellStr); } catch(e){} }
      }

      // 檢查該使用者是否與此專案有關聯 (PM 或 關卡執行者)
      let isInvolved = (pmName === userName);
      let assignedSteps = [];
      
      if (wfData.deliverables) {
        wfData.deliverables.forEach(d => {
          if (d.workflow) {
            d.workflow.forEach(s => {
              if (s.assignee && s.assignee.toLowerCase() === userName) {
                isInvolved = true;
                if (s.status === 'In Progress') assignedSteps.push({ stepName: s.name, deadline: s.keyDate || s.deadline });
              }
            });
          }
        });
      }

      if (!isInvolved) continue; // 如果與我無關，跳過這個專案

      // === A. 狀態型通知 (Live Status) ===
      assignedSteps.forEach(s => {
        let deadline = s.deadline || todayStr;
        let daysLeft = Math.ceil((new Date(deadline) - new Date(todayStr)) / 86400000);
        
        if (daysLeft < 0) {
          notifications.push({ type: 'OVERDUE', color: 'danger', icon: 'fa-exclamation-triangle', time: new Date().toISOString(),
            title: `[${jobNumber}] 任務已逾期！`, message: `你的任務 ${s.stepName} 已經逾期，請盡速處理。` });
        } else if (daysLeft <= 1) {
          notifications.push({ type: 'DUE_SOON', color: 'warning', icon: 'fa-hourglass-half', time: new Date().toISOString(),
            title: `[${jobNumber}] 任務即將到期`, message: `你的任務 ${s.stepName} 需於今天/明天內完成。` });
        }
      });

      // === B. 事件型通知 (Event Logs) ===
      logs.forEach(log => {
        let logDate = new Date(log.timestamp.replace(' ', 'T') + ':00');
        if (isNaN(logDate) || logDate < sevenDaysAgo) return; // 只抓最近一週

        const action = String(log.action || '').toLowerCase();
        const details = String(log.details || '').toLowerCase();
        const logUser = String(log.user || '').toLowerCase();

        // 不要通知自己做過的事 (除非是重要的狀態改變)
        if (logUser === userName && !action.includes('nudge')) return;

        let notif = null;

        if (action.includes('dispatch task')) {
          if (details.includes(`指派給 [${userName}]`)) {
            notif = { type: 'ASSIGNED', color: 'primary', icon: 'fa-inbox', title: `[${jobNumber}] 新任務派發`, message: `${log.user} 指派了新任務給你。` };
          } else if (details.includes(`由 [${userName}] 更換為`)) {
            notif = { type: 'REASSIGNED', color: 'secondary', icon: 'fa-random', title: `[${jobNumber}] 任務轉交`, message: `你原本負責的任務已更換由其他同事接手。` };
          }
        } 
        else if (action.includes('insert step')) {
          notif = { type: 'SCOPE_CHANGE', color: 'info', icon: 'fa-plus-circle', title: `[${jobNumber}] 流程變更`, message: `${log.user} 在專案中插入了新的關卡步驟。` };
        }
        else if (action.includes('revision') || action.includes('rollback')) {
          notif = { type: 'REJECTED', color: 'danger', icon: 'fa-undo', title: `[${jobNumber}] 專案退回修改`, message: `${log.user} 提出修改要求或退回了關卡：${log.details}` };
        }
        else if (action.includes('nudge')) {
          notif = { type: 'NUDGE', color: 'danger', icon: 'fa-bullhorn', title: `[${jobNumber}] 📢 催辦提醒！`, message: `${log.user} 對你負責的關卡發出了催辦，請盡速查看！` };
        }
        else if (action.includes('submit step') && pmName === userName) {
          // PM 專屬：下屬完成關卡
          notif = { type: 'STEP_DONE', color: 'success', icon: 'fa-check-circle', title: `[${jobNumber}] 關卡完成`, message: `${log.user} 已完成一項關卡，專案繼續推進。` };
        }
        else if (action.includes('pause project')) {
          notif = { type: 'PAUSED', color: 'warning', icon: 'fa-pause-circle', title: `[${jobNumber}] 專案暫停`, message: `專案已被 ${log.user} 暫停。` };
        }

        if (notif) {
          notif.time = logDate.toISOString();
          notif.rawTime = log.timestamp;
          notifications.push(notif);
        }
      });
    }

    // 依照時間由新到舊排序
    notifications.sort((a, b) => new Date(b.time).getTime() - new Date(a.time).getTime());
    
    // 最多只回傳最新的 30 則通知
    return { success: true, data: notifications.slice(0, 30) };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function testSendEmailNow() {
  // 🔽 把這裡換成你正在檢查的真實信箱！
  var myRealEmail = "raylo@hk01.com"; 

  sendSystemEmail(
    myRealEmail, 
    "🧪 系統發信功能測試", 
    "測試成功！", 
    "如果你看到這封信，代表發信引擎已經成功獲得授權並運作正常！"
  );
}

function sendSystemEmail(toUser, subject, title, message) {
  console.log(`[發信模組] 接收到目標參數: [${toUser}]`);
  if (!toUser || toUser === '未指派') return '無負責人';
  
  let emailAddress = toUser;
  
  if (!emailAddress.includes('@')) {
    try {
      const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
      if (userSheet) {
        const data = userSheet.getDataRange().getValues();
        const headers = data[0].map(h => String(h || '').trim().toLowerCase());
        const idxEmail = headers.findIndex(h => h === 'email' || h.includes('email'));
        const idxName = headers.findIndex(h => h === '姓名' || h === 'name');
        
        if (idxEmail >= 0 && idxName >= 0) {
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][idxName]).trim() === String(toUser).trim()) {
              emailAddress = String(data[i][idxEmail]).trim();
              console.log(`[發信模組] 查表成功！將 [${toUser}] 轉換為 [${emailAddress}]`);
              break;
            }
          }
        }
      }
    } catch(e) {
      console.log(`[發信模組] 查表發生錯誤: ${e.message}`);
    }
    
    if (!emailAddress.includes('@')) {
      emailAddress = emailAddress.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() + '@hk01.com';
      console.log(`[發信模組] 查表失敗，強制轉為預設後綴: [${emailAddress}]`);
    }
  }

  let sysUrl = ScriptApp.getService().getUrl(); 
  let htmlBody = `
    <div style="font-family: sans-serif; max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #005088; color: white; padding: 15px 20px; font-size: 18px; font-weight: bold;">HK01 PMO 系統通知</div>
      <div style="padding: 20px; color: #333; line-height: 1.6;">
        <h3 style="color: #005088; margin-top: 0;">${title}</h3>
        <p>${message}</p>
        <a href="${sysUrl}" style="display: inline-block; margin-top: 20px; padding: 10px 20px; background-color: #0dcaf0; color: #000; text-decoration: none; border-radius: 5px; font-weight: bold;">登入系統處理</a>
      </div>
    </div>
  `;
  try {
    GmailApp.sendEmail(emailAddress, subject, "", { name: 'HK01 PMO 系統', htmlBody: htmlBody });
    console.log(`[發信模組] 信件已成功呼叫 GmailApp 寄出至: [${emailAddress}]`);
    return emailAddress; 
  } catch(e) { 
    console.error(`[發信模組] 呼叫 GmailApp 失敗: ${e.message}`);
    return "錯誤: " + e.message; 
  }
}

// ==========================================
// 💡 全量 Firebase 廣播與快取推送引擎 (秒讀基礎)
// ==========================================
function notifyFirebaseUpdate(jobNumber, userEmail, userName, isGlobalEvent) {
  try {
    const cleanJobNum = String(jobNumber || 'GLOBAL').split('-P')[0];
    const firebaseUrl = "https://hk01-pmo-realtime-default-rtdb.asia-southeast1.firebasedatabase.app/";
    
    const nowStamp = new Date().getTime();
    const payload = JSON.stringify({
      timestamp: nowStamp,
      updatedBy: userEmail || 'System',
      updatedByName: userName || 'System'
    });

    if (cleanJobNum !== 'GLOBAL') {
      // 1. 廣播時間戳記
      UrlFetchApp.fetch(firebaseUrl + "projects/" + cleanJobNum + ".json", {
        method: "patch", contentType: "application/json", payload: payload, muteHttpExceptions: true
      });

      // 2. 🚀 將專案完整內容寫入 Firebase 快取 (存於 project_cache/ 下)
      try {
        let freshProject = api_getProjectWorkflow(cleanJobNum);
        if (freshProject && freshProject.success) {
          UrlFetchApp.fetch(firebaseUrl + "project_cache/" + cleanJobNum + ".json", {
            method: "put",
            contentType: "application/json",
            payload: JSON.stringify(freshProject.data),
            muteHttpExceptions: true
          });
        }
      } catch(e) {}
    }

    if (isGlobalEvent) {
      UrlFetchApp.fetch(firebaseUrl + "global_events.json", {
        method: "patch", contentType: "application/json", payload: payload, muteHttpExceptions: true
      });
    }
  } catch(e) {
    console.error("Firebase 廣播失敗:", e.message);
  }
}

// ==========================================
// 💡 後端時間格式化輔助函數 (解決 formatDurationText undefined 報錯)
// ==========================================
function formatDurationText(ms) {
  if (!ms || ms <= 0) return '0 分鐘';
  var totalMins = Math.floor(ms / (1000 * 60));
  var mins = totalMins % 60;
  var totalHours = Math.floor(totalMins / 60);
  var hours = totalHours % 24;
  var days = Math.floor(totalHours / 24);

  var result = '';
  if (days > 0) result += days + ' 天 ';
  if (hours > 0) result += hours + ' 小時 ';
  result += mins + ' 分鐘';
  return result.trim();
}

// ==========================================
// 💡 全域日期格式化輔助函數 (統一輸出 YYYY-MM-DD)
// ==========================================
function cleanYMD(raw) {
  if (!raw) return '未設定';
  if (raw instanceof Date) return Utilities.formatDate(raw, "GMT+8", "yyyy-MM-dd");
  let str = String(raw).trim();
  if (str.includes('GMT') || str.includes('Standard Time') || str.includes('T')) {
    let d = new Date(str);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd");
  }
  return str.split('T')[0].split(' ')[0];
}
