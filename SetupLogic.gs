/**
 * 📄 SetupLogic.gs
 * 系統初始化 (Setup)、重置 (Reset) 與數據備份 (Backup) 工具
 */

// ==========================================
// 1. 一鍵初始化系統 (setupSystem)
// ==========================================
/**
 * 自動建立試算表結構、預填資料、建立 Drive 資料夾與排程
 */
function setupSystem() {
  const logger = [];
  function log(msg) {
    Logger.log(msg);
    logger.push(`[${new Date().toLocaleTimeString()}] ${msg}`);
  }

  log('🚀 開始執行全系統一鍵 Setup...');

  try {
    // 1.1 初始化所有工作表結構 (呼叫 Config.gs 中的函數)
    log('1/6 建立工作表 (Sheets) 與標題列...');
    initializeSheets();
    log('✅ 工作表建立成功。');

    // 1.2 預寫入 Config 設定檔
    log('2/6 寫入預設 Config 設定...');
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
    const defaultConfig = [
      ['SLA_EDITORIAL_DAYS', '3', '文字/採訪預設工作天數'],
      ['SLA_DESIGN_DAYS', '1', '設計/美術預設工作天數'],
      ['MAX_VERSION', '5', '觸發警告的最高修改版本號'],
      ['DRIVE_PARENT_FOLDER_NAME', 'HK01_PMO_Projects', '雲端硬碟專案主資料夾名稱']
    ];
    // 如果只有標題列，寫入預設值
    if (configSheet.getLastRow() <= 1) {
      configSheet.getRange(2, 1, defaultConfig.length, 3).setValues(defaultConfig);
      log('✅ 預設 Config 寫入完成。');
    } else {
      log('ℹ️ Config 表已有內容，跳過預設值寫入。');
    }

    // 1.3 預寫入 35 個行業分類到 Industries Sheet
    log('3/6 寫入 35 個行業分類...');
    const industriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.INDUSTRIES);
    if (industriesSheet.getLastRow() <= 1) {
      const industries = [
        ['地產 (Real Estate)'], ['金融 (Finance)'], ['銀行 (Banking)'], ['保險 (Insurance)'],
        ['科技 (Technology)'], ['電子商務 (E-commerce)'], ['零售 (Retail)'], ['飲食 (F&B)'],
        ['旅遊 (Travel)'], ['酒店 (Hospitality)'], ['汽車 (Automotive)'], ['美容 (Beauty)'],
        ['個人護理 (Personal Care)'], ['時尚/服飾 (Fashion)'], ['奢侈品 (Luxury)'], ['醫療/健康 (Healthcare)'],
        ['製藥 (Pharmaceutical)'], ['教育 (Education)'], ['政府/公營機構 (Government)'], ['非牟利組織 (NGO)'],
        ['快消品 (FMCG)'], ['母嬰/兒童 (Baby & Kids)'], ['寵物 (Pets)'], ['電玩/遊戲 (Gaming)'],
        ['娛樂 (Entertainment)'], ['體育/健身 (Sports)'], ['家電/數碼 (Electronics)'], ['家居/裝修 (Home)'],
        ['物流/運輸 (Logistics)'], ['電訊 (Telecom)'], ['專業服務 (Professional Services)'], ['能源/環保 (Energy)'],
        ['文化/藝術 (Culture)'], ['媒體/廣告 (Media)'], ['其他 (Others)']
      ];
      industriesSheet.getRange(2, 1, industries.length, 1).setValues(industries);
      log('✅ 35 個行業分類寫入完成。');
    } else {
      log('ℹ️ Industries 表已有內容，跳過寫入。');
    }

    // 1.4 預寫入所有 PRD 提到的預設使用者到 Users Sheet
    log('4/6 寫入預設使用者資料庫...');
    const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.USERS);
    if (usersSheet.getLastRow() <= 1) {
      // 依據 PRD 角色與部門進行編排
      const defaultUsers = [
        // Name, Email, Role, Department, PhoneNumber
        ['Ray', 'ray@hk01.com', 'Admin', 'Management', '85290000001'],
        ['Wincy', 'wincy@hk01.com', 'PM', 'PM', '85290000002'],
        ['Yannes', 'yannes@hk01.com', 'PM', 'PM', '85290000003'],
        ['Wing', 'wing@hk01.com', 'Team Head', 'Editorial', '85290000004'],
        ['Ming', 'ming@hk01.com', 'Editor', 'Editorial', '85290000005'],
        ['Wicky', 'wicky@hk01.com', 'Editor', 'Editorial', '85290000006'],
        ['Man', 'man@hk01.com', 'Editor', 'Editorial', '85290000007'],
        ['Winnie', 'winnie@hk01.com', 'Team Head', 'Creative', '85290000008'],
        ['Enson', 'enson@hk01.com', 'Designer', 'Creative', '85290000009'],
        ['Vince', 'vince@hk01.com', 'Designer', 'Creative', '85290000010'],
        ['Hung', 'hung@hk01.com', 'Designer', 'Creative', '85290000011'],
        ['Jacky', 'jacky@hk01.com', 'Sales', 'Sales', '85290000012'],
        ['Carman', 'carman@hk01.com', 'Sales', 'Sales', '85290000013'],
        ['Yan', 'yan@hk01.com', 'Sales', 'Sales', '85290000014'],
        ['Martini', 'martini@hk01.com', 'Management', 'Management', '85290000015']
      ];
      usersSheet.getRange(2, 1, defaultUsers.length, 5).setValues(defaultUsers);
      log('✅ 15 位預設使用者寫入完成。');
    } else {
      log('ℹ️ Users 表已有內容，跳過寫入。');
    }

    // 1.5 檢查/建立 Google Drive 父資料夾
    log('5/6 檢查與建立 Google Drive 父資料夾...');
    const parentFolder = getParentFolder();
    log(`✅ Google Drive 資料夾已就緒 (Folder ID: ${parentFolder.getId()})`);

    // 1.6 安裝排程觸發器 (Time Triggers)
    log('6/6 安裝背景排程觸發條件 (Triggers)...');
    setupTriggers();
    log('✅ 每日排程檢測安裝完成。');

    // 記錄初始化 Log
    logAction('SYSTEM_SETUP', 'ALL', '全系統一鍵 Setup 完成');

    log('🎉🎉🎉 全系統初始化成功！您可以開始使用 HK01 PMO 系統。');
    
    // 彈出成功提示
    //SpreadsheetApp.getUi().alert('系統初始化完成', logger.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    log('❌ 初始化失敗: ' + error.message);
    //SpreadsheetApp.getUi().alert('初始化失敗', '錯誤訊息: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// 2. 系統重置 (resetSystem - 附二次確認防呆)
// ==========================================
/**
 * 清空專案與任務資料，但保留表格欄位結構與 Config/Users 基礎設定
 */
function resetSystem() {
  const ui = SpreadsheetApp.getUi();
  
  // 第一重警告
  const response1 = ui.alert(
    '⚠️ 嚴重警告：系統重置',
    '此操作將會「刪除所有專案 (Projects)、任務 (Tasks) 與審計日誌 (AuditLog)」！\n請問您確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    ui.alert('操作已取消，資料未受任何影響。');
    return;
  }

  // 第二重警告 (防止誤觸)
  const response2 = ui.prompt(
    '🚨 二次確認 (Safety Check)',
    '這是最後警告！請在下方輸入框寫入 "RESET" 字母 (全大寫) 以確認清空系統：',
    ui.ButtonSet.OK_CANCEL
  );

  if (response2.getSelectedButton() === ui.Button.OK && response2.getResponseText().trim() === 'RESET') {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // 需要清空數據的工作表
      const sheetsToClear = [SHEET_NAMES.PROJECTS, SHEET_NAMES.TASKS, SHEET_NAMES.AUDIT_LOG];
      
      sheetsToClear.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet && sheet.getLastRow() > 1) {
          // 清除第 2 列之後的所有內容與格式
          sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
        }
      });

      logAction('SYSTEM_RESET', 'ALL', '使用者手動觸發系統重置，已清空 Projects, Tasks 與 AuditLog');
      ui.alert('✅ 重置成功', '系統動態資料庫已清空，基礎設定 (Users, Config) 已保留。', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('❌ 重置失敗: ' + e.message);
    }
  } else {
    ui.alert('驗證碼不匹配，重置操作已安全取消。');
  }
}

// ==========================================
// 3. 數據備份匯出 (exportAllData)
// ==========================================
/**
 * 將目前整個試算表複製一份到 Google Drive 作為備份，並提供下載連結
 */
function exportAllData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dateStr = formatDate(new Date()).replace(/-/g, '');
    const backupName = `HK01_PMO_Backup_${dateStr}_${new Date().getTime()}`;

    // 在同一個 Google Drive 建立試算表副本
    const file = DriveApp.getFileById(ss.getId());
    const backupFile = file.makeCopy(backupName);

    logAction('SYSTEM_BACKUP', 'ALL', `匯出系統備份檔: ${backupName}`);

    const downloadUrl = backupFile.getUrl();
    
    ui.alert(
      '📦 備份完成！',
      `數據已成功複製備份至您的 Google Drive！\n\n備份檔名稱：${backupName}\n\n您可以開啟以下網址存取備份檔：\n${downloadUrl}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ 匯出備份失敗: ' + e.message);
  }
}

// ==========================================
// 4. Google Drive 安全取得/建立資料夾輔助函數
// ==========================================
/**
 * 取得或建立專案的 Google Drive 主資料夾
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getParentFolder() {
  const folderName = 'HK01_PMO_Projects';
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    // 若不存在則自動建立主資料夾
    return DriveApp.createFolder(folderName);
  }
}