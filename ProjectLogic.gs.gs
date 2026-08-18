// ==========================================
// 4. 專案 CRUD 核心功能
// ==========================================

/**
 * 新增專案
 * @param {Object} data - 專案基礎資料
 * @returns {string} - 新建立的 jobNumber
 */
function createProject(data) {
  try {
    const jobNumber = generateJobNumber();
    const submissionDate = new Date();
    
    // 組合新專案資料
    const newProject = {
      ...data,
      jobNumber: jobNumber,
      submissionDate: submissionDate,
      status: 'Not Started',
      createdAt: submissionDate,
      updatedAt: submissionDate
    };

    // 寫入資料庫 (appendRow 會自動處理數字欄位補 0 的邏輯)
    appendRow(SHEET_NAMES.PROJECTS, newProject);
    
    // 記錄 Log (呼叫 SheetHelper.gs 裡的 logAction)
    logAction('CREATE_PROJECT', jobNumber, `建立新專案: ${data.clientName || '未命名客戶'}`);

    return jobNumber;
  } catch (error) {
    throw new Error('新增專案失敗: ' + error.message);
  }
}

/**
 * 讀取單個專案資料
 * @param {string} jobNumber 
 * @returns {Object}
 */
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

/**
 * 更新專案資料
 * @param {string} jobNumber 
 * @param {Object} updates - 要更新的欄位 (例如 { status: 'In Progress' })
 * @returns {boolean}
 */
function updateProject(jobNumber, updates) {
  try {
    updates.updatedAt = new Date();
    const success = updateRow(SHEET_NAMES.PROJECTS, 'jobNumber', jobNumber, updates);
    if (!success) throw new Error('更新失敗，找不到對應的 Job Number');
    
    // 篩選出更新的 Key 來寫 Log
    const updatedKeys = Object.keys(updates).join(', ');
    logAction('UPDATE_PROJECT', jobNumber, `更新欄位: ${updatedKeys}`);
    
    return true;
  } catch (error) {
    throw new Error('更新專案失敗: ' + error.message);
  }
}

/**
 * 啟動專案 (PM 點擊 Start)
 * @param {string} jobNumber 
 */
function startProject(jobNumber) {
  try {
    const project = getProject(jobNumber);
    
    // 1. 防呆警告：檢查 Launch Date
    if (!project.launchDate || project.launchDate === '') {
      throw new Error('【防呆警告】必須填寫 Launch Date 才能啟動專案！');
    }

    // 2. 自動建立 Google Drive 資料夾 (目前先寫 Mock，等 DriveLogic 完成後對接)
    // const folderUrl = createProjectFolder(jobNumber, project.clientName, project.productName);
    const folderUrl = `https://drive.google.com/mock_folder/${jobNumber}`; 
    
    // 3. 更新狀態與 URL
    updateProject(jobNumber, {
      status: 'Waiting for PIC',
      googleDriveFolderUrl: folderUrl
    });
    logAction('START_PROJECT', jobNumber, '專案正式啟動，狀態改為 Waiting for PIC');

    // 4. 根據項目自動發送通知 (目前先用 Logger 印出，等 NotificationLogic 完成後對接)
    if (project.advertorialArticle > 0 || project.booklet > 0 || project.textJobType === 'Advertorial') {
       Logger.log(`🔔 [通知] 已自動發送任務提醒給 Editor Team`);
    }
    if (project.video > 0 || project.bannerAd > 0) {
       Logger.log(`🔔 [通知] 已自動發送任務提醒給 Creative/Art Team`);
    }

    return folderUrl;
  } catch (error) {
    throw new Error('啟動專案失敗: ' + error.message);
  }
}

/**
 * 列出專案 (支援權限過濾與條件篩選)
 * @param {Object} filter - 篩選條件 (如 { status: 'Active' })
 * @returns {Array<Object>}
 */
function listProjects(filter = {}) {
  try {
    let projects = getSheetData(SHEET_NAMES.PROJECTS);
    
    // 1. RBAC 權限自動過濾 (只保留有權限看的專案)
    projects = projects.filter(p => canViewProject(p.jobNumber));

    // 2. 應用自訂篩選條件
    if (filter.status) projects = projects.filter(p => p.status === filter.status);
    if (filter.pmName) projects = projects.filter(p => p.pmName === filter.pmName);
    
    return projects;
  } catch (error) {
    throw new Error('列出專案失敗: ' + error.message);
  }
}

// ==========================================
// 5. 測試 CRUD 流程
// ==========================================
function testProjectCRUD() {
  try {
    Logger.log('--- 1. 測試建立專案 ---');
    const newJobNumber = createProject({
      clientName: 'HK01 Test Client',
      productName: 'Summer Campaign',
      pmName: 'Wincy',
      launchDate: '2026-08-01',
      advertorialArticle: 2 // 故意加入 Advertorial 項目測試通知
    });
    Logger.log(`✅ 專案建立成功！Job Number: ${newJobNumber}`);
    
    Logger.log('--- 2. 測試讀取專案 ---');
    const project = getProject(newJobNumber);
    Logger.log(`✅ 讀取成功！客戶名稱: ${project.clientName}, 狀態: ${project.status}`);

    Logger.log('--- 3. 測試啟動專案 (Start Project) ---');
    const url = startProject(newJobNumber);
    Logger.log(`✅ 專案啟動成功！綁定的 Drive 網址: ${url}`);

    Logger.log('--- 4. 測試權限過濾與列表 ---');
    const myList = listProjects();
    Logger.log(`✅ 您目前有權限查看的專案數量: ${myList.length}`);
    
  } catch (error) {
    Logger.log('❌ 測試失敗: ' + error.message);
  }
}