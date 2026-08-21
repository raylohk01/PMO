/**
 * 📄 ProjectLogic.gs
 * 專案核心邏輯、自動編號生成與 CRUD 處理 (整合修復版)
 */

// ==========================================
// 1. 輔助函數：解析 Job Number 年份
// ==========================================
function getYearFromJobNumber(jobNumber) {
  if (!jobNumber || typeof jobNumber !== 'string') return null;
  const match = jobNumber.match(/^A(\d{2})-\d{4,}$/);
  return match ? match[1] : null;
}

// ==========================================
// 2. 主函數：自動生成流水號 Job Number (如 A26-1001)
// ==========================================
function generateJobNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  if (!sheet) throw new Error('找不到 Projects 表格，請確認是否已經執行 initializeSheets。');

  const currentYear = new Date().getFullYear().toString().slice(-2);
  const targetPrefix = `A${currentYear}-`;
  const data = sheet.getDataRange().getValues();
  let maxSequence = 1000;

  for (let i = 1; i < data.length; i++) {
    const existingJobNumber = data[i][0];
    if (existingJobNumber && typeof existingJobNumber === 'string' && existingJobNumber.startsWith(targetPrefix)) {
      const sequenceStr = existingJobNumber.split('-')[1];
      const sequenceNum = parseInt(sequenceStr, 10);
      if (!isNaN(sequenceNum) && sequenceNum > maxSequence) {
        maxSequence = sequenceNum;
      }
    }
  }
  return `${targetPrefix}${maxSequence + 1}`;
}

// ==========================================
// 3. 專案 CRUD 核心功能 (含範本工作流自動對接機制)
// ==========================================

/**
 * 新增專案 (自動載入 WorkflowTemplates 範本)
 */
function createProject(clientName, productNameInput, salesPerson, productType) {
  try {
    let data = {};
    if (typeof clientName === 'object' && clientName !== null) {
      data = clientName;
    } else {
      data = {
        clientName: clientName || '未命名客戶',
        productName: productNameInput || '{"deliverables":[]}',
        salesPerson: salesPerson || '未指定',
        productType: productType || '標準 Advertorial'
      };
    }

    const jobNumber = generateJobNumber();
    const submissionDate = new Date();

    // 💡 徹底防護：不論前端傳來的是空字串、空物件、或是 "deliverables":[] 都能精準抓包並自動帶入範本！
    var deliverablesData = [];
    try {
      if (typeof data.productName === 'string') {
        deliverablesData = JSON.parse(data.productName).deliverables || [];
      } else if (data.deliverables) {
        deliverablesData = data.deliverables;
      } else if (data.productName && data.productName.deliverables) {
        deliverablesData = data.productName.deliverables;
      }
    } catch(e) {}

    // 只要發現 deliverables 是空的，一律強制去 WorkflowTemplates 抓取「標準 Advertorial」寫入！
    if (!deliverablesData || deliverablesData.length === 0) {
      var templateName = data.productType || data.templateName || '標準 Advertorial';
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var templateSheet = ss.getSheetByName('WorkflowTemplates');
      var workflowSteps = [];

      if (templateSheet) {
        var templatesData = templateSheet.getDataRange().getValues();
        for (var i = 1; i < templatesData.length; i++) {
          if (String(templatesData[i][0]).trim() === templateName) {
            try { workflowSteps = JSON.parse(templatesData[i][1]); } catch(e){}
            break;
          }
        }
      }

      deliverablesData = [{
        id: 'deliv_1_' + Date.now(),
        name: data.deliverableName || '篇章 / 任務 1',
        type: templateName,
        status: 'Pending Start',
        workflow: workflowSteps
      }];
    }

    var finalProductName = JSON.stringify({ deliverables: deliverablesData });

    // 組合新專案資料結構
    const newProject = {
      jobNumber: jobNumber,
      submissionDate: submissionDate,
      launchDate: data.launchDate || data.deadline || submissionDate,
      status: 'In Progress',
      clientName: data.clientName || data.client || '未命名客戶',
      productName: finalProductName,
      salesPerson: data.salesPerson || data.sales || '未指定',
      pmName: data.pmName || data.pm || '',
      createdAt: submissionDate,
      updatedAt: submissionDate
    };

    if (typeof appendRow === 'function' && typeof SHEET_NAMES !== 'undefined') {
      appendRow(SHEET_NAMES.PROJECTS, newProject);
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const pSheet = ss.getSheetByName('Projects');
      pSheet.appendRow([
        newProject.jobNumber,
        newProject.submissionDate,
        newProject.launchDate,
        '',
        newProject.status,
        '',
        newProject.clientName,
        newProject.productName,
        '[]',
        '',
        newProject.salesPerson,
        newProject.pmName
      ]);
    }
    
    if (typeof logAction === 'function') {
      logAction('CREATE_PROJECT', jobNumber, `建立了專案: ${newProject.clientName}`);
    }

    return jobNumber;
  } catch (error) {
    throw new Error('新增專案失敗: ' + error.message);
  }
}

function getProject(jobNumber) {
  try {
    const projects = getSheetData(SHEET_NAMES.PROJECTS);
    const project = projects.find(p => p.jobNumber === jobNumber);
    if (!project) throw new Error('找不到專案: ' + jobNumber);
    return project;
  } catch (error) {
    throw new Error('讀取專案失敗: ' + error.message);
  }
}

function updateProject(jobNumber, updates) {
  try {
    updates.updatedAt = new Date();
    const success = updateRow(SHEET_NAMES.PROJECTS, 'jobNumber', jobNumber, updates);
    if (!success) throw new Error('更新失敗，找不到對應的 Job Number');
    
    const updatedKeys = Object.keys(updates).join(', ');
    logAction('UPDATE_PROJECT', jobNumber, `更新欄位: ${updatedKeys}`);
    return true;
  } catch (error) {
    throw new Error('更新專案失敗: ' + error.message);
  }
}

function startProject(jobNumber) {
  try {
    const project = getProject(jobNumber);
    if (!project.launchDate || project.launchDate === '') {
      throw new Error('【防呆警告】必須填寫 Launch Date 才能啟動專案！');
    }

    const folderUrl = `https://drive.google.com/mock_folder/${jobNumber}`; 
    updateProject(jobNumber, {
      status: 'Waiting for PIC',
      googleDriveFolderUrl: folderUrl
    });
    logAction('START_PROJECT', jobNumber, '專案正式啟動，狀態改為 Waiting for PIC');
    return folderUrl;
  } catch (error) {
    throw new Error('啟動專案失敗: ' + error.message);
  }
}

function listProjects(filter = {}) {
  try {
    let projects = getSheetData(SHEET_NAMES.PROJECTS);
    if (typeof canViewProject === 'function') {
      projects = projects.filter(p => canViewProject(p.jobNumber));
    }
    if (filter.status) projects = projects.filter(p => p.status === filter.status);
    if (filter.pmName) projects = projects.filter(p => p.pmName === filter.pmName);
    return projects;
  } catch (error) {
    throw new Error('列出專案失敗: ' + error.message);
  }
}