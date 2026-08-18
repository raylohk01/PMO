/**
 * 📄 DriveLogic.gs
 * Google Drive 自動化與資料夾管理
 */

// ==========================================
// 1. 輔助函數
// ==========================================

/**
 * 從 Config Sheet 取得父資料夾 ID
 * @returns {string} - 父資料夾 ID
 */
function getParentFolderId() {
  const configData = getSheetData(SHEET_NAMES.CONFIG);
  const folderConfig = configData.find(c => c['設定名稱'] === 'DRIVE_PARENT_ID');
  
  if (!folderConfig || !folderConfig['設定值']) {
    throw new Error('請先在 Config 表中設定 DRIVE_PARENT_ID (共用雲端硬碟的父資料夾 ID)。');
  }
  return folderConfig['設定值'];
}

/**
 * 尋找指定專案的資料夾 ID
 * @param {string} jobNumber 
 * @returns {string} - 資料夾 ID
 */
function getProjectFolderByIdOrSearch(jobNumber) {
  // 先嘗試從 Projects Sheet 獲取已經記錄的 URL
  try {
    const project = getProject(jobNumber);
    if (project && project.googleDriveFolderUrl) {
      // 擷取 URL 中的 ID 片段
      const match = project.googleDriveFolderUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
      if (match && match[1]) return match[1];
    }
  } catch (e) {
    // 若尚未記錄或發生錯誤，轉由直接搜尋 Drive
  }
  
  // 在 Drive 中透過名稱搜尋資料夾
  const folders = DriveApp.searchFolders(`title contains '${jobNumber}_' and trashed = false`);
  if (folders.hasNext()) {
    return folders.next().getId();
  }
  throw new Error(`找不到專案 ${jobNumber} 的專屬資料夾`);
}

// ==========================================
// 2. 核心功能
// ==========================================

/**
 * 自動建立專案資料夾
 * @param {string} jobNumber 
 * @param {string} clientName 
 * @param {string} productName 
 * @returns {string} - 新建立的資料夾 URL
 */
function createProjectFolder(jobNumber, clientName, productName) {
  try {
    const parentId = getParentFolderId();
    const parentFolder = DriveApp.getFolderById(parentId);
    
    // 檔名消毒：避免出現影響路徑的特殊符號 (如斜線)
    const safeClientName = clientName ? clientName.replace(/\//g, '-') : 'UnknownClient';
    const safeProductName = productName ? productName.replace(/\//g, '-') : 'UnknownProduct';
    const folderName = `${jobNumber}_${safeClientName}_${safeProductName}`;
    
    // 建立資料夾
    const newFolder = parentFolder.createFolder(folderName);
    
    // 處理權限：設定網域內知道連結的使用者皆可編輯，避免同事存取被阻擋
    // (註：部分 Shared Drive 會強制鎖定權限，此處以 try-catch 包覆避免報錯中斷)
    try {
      newFolder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (permError) {
      Logger.log(`設定權限時發生警告 (共用雲端硬碟可能已強制繼承權限): ${permError.message}`);
    }

    const folderUrl = newFolder.getUrl();
    Logger.log(`✅ 成功建立資料夾: ${folderName}`);
    
    return folderUrl;
  } catch (error) {
    throw new Error('建立專案資料夾失敗: ' + error.message);
  }
}

/**
 * 上傳檔案到專案資料夾
 * @param {string} jobNumber 
 * @param {string} fileName 
 * @param {GoogleAppsScript.Base.Blob} fileBlob 
 * @returns {Object} - 上傳結果與檔案 URL
 */
function uploadFileToProjectFolder(jobNumber, fileName, fileBlob) {
  try {
    const folderId = getProjectFolderByIdOrSearch(jobNumber);
    const folder = DriveApp.getFolderById(folderId);
    
    const newFile = folder.createFile(fileBlob);
    newFile.setName(fileName);
    
    return {
      success: true,
      fileUrl: newFile.getUrl(),
      fileName: fileName
    };
  } catch (error) {
    throw new Error('檔案上傳失敗: ' + error.message);
  }
}

/**
 * 列出專案資料夾內的所有檔案
 * @param {string} jobNumber 
 * @returns {Array<Object>} - 檔案清單陣列
 */
function listProjectFiles(jobNumber) {
  try {
    const folderId = getProjectFolderByIdOrSearch(jobNumber);
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    const fileList = [];
    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        name: file.getName(),
        url: file.getUrl(),
        dateCreated: file.getDateCreated(),
        size: file.getSize()
      });
    }
    
    return fileList;
  } catch (error) {
    throw new Error('列出檔案失敗: ' + error.message);
  }
}

// ==========================================
// 3. 測試程式碼
// ==========================================
function testDriveLogic() {
  // 測試前準備：自動在 Config 寫入 Google Drive 根目錄 (root) 作為暫時的測試用父資料夾
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  const configData = configSheet.getDataRange().getValues();
  let hasDriveConfig = false;
  
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === 'DRIVE_PARENT_ID') {
      configSheet.getRange(i + 1, 2).setValue('root'); // 強制設定為 root 確保測試能跑
      hasDriveConfig = true;
      break;
    }
  }
  if (!hasDriveConfig) {
    configSheet.appendRow(['DRIVE_PARENT_ID', 'root']);
  }
  
  try {
    Logger.log('--- 1. 建立資料夾測試 ---');
    const folderUrl = createProjectFolder('A26-9999', 'HK01', 'Drive_Test_Campaign');
    Logger.log(`建立成功！URL: ${folderUrl}`);
    
    Logger.log('--- 2. 上傳檔案測試 ---');
    // 動態產生一個測試用的純文字檔 (Blob)
    const blob = Utilities.newBlob('這是一份透過 GAS 自動上傳的測試文件。', 'text/plain', 'Test_Document.txt');
    const uploadRes = uploadFileToProjectFolder('A26-9999', 'Test_Document.txt', blob);
    Logger.log(`上傳成功！檔案 URL: ${uploadRes.fileUrl}`);
    
    Logger.log('--- 3. 列表檔案測試 ---');
    const files = listProjectFiles('A26-9999');
    Logger.log(`共找到 ${files.length} 個檔案：`);
    files.forEach(f => Logger.log(`- ${f.name} (${f.url})`));
    
  } catch (e) {
    Logger.log('❌ 測試發生錯誤: ' + e.message);
  }
}