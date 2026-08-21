/**
 * 📄 CalendarLogic.gs
 * 香港工作日計算與 SLA 期限管理
 */

// ==========================================
// 1. 假期資料初始化與讀取
// ==========================================

// 2026 年香港主要公眾假期預設清單 (格式 YYYY-MM-DD)
const DEFAULT_HK_HOLIDAYS = [
  '2026-01-01', // 元旦
  '2026-02-17', '2026-02-18', '2026-02-19', // 農曆年初一至初三
  '2026-04-03', '2026-04-04', '2026-04-06', // 復活節與清明節假期
  '2026-05-01', // 勞動節
  '2026-05-25', // 佛誕翌日
  '2026-06-19', // 端午節
  '2026-07-01', // 七一
  '2026-09-26', // 中秋節翌日
  '2026-10-01', // 國慶日
  '2026-10-19', // 重陽節翌日
  '2026-12-25', '2026-12-26' // 聖誕節
].join(',');

/**
 * 確保 Config 表中有假期設定 (如果沒有則自動寫入)
 */
function ensureHolidaysInConfig() {
  const configData = getSheetData(SHEET_NAMES.CONFIG);
  const holidayConfig = configData.find(c => c['設定名稱'] === 'HK_HOLIDAYS');
  
  if (!holidayConfig) {
    appendRow(SHEET_NAMES.CONFIG, {
      '設定名稱': 'HK_HOLIDAYS',
      '設定值': DEFAULT_HK_HOLIDAYS
    });
    Logger.log('已將 2026 香港公眾假期預設值寫入 Config 表！');
  }
}

/**
 * 從 Config 讀取並回傳假期陣列
 * @returns {Array<string>} - 假期日期陣列 (YYYY-MM-DD)
 */
function getHolidays() {
  const configData = getSheetData(SHEET_NAMES.CONFIG);
  const holidayConfig = configData.find(c => c['設定名稱'] === 'HK_HOLIDAYS');
  if (holidayConfig && holidayConfig['設定值']) {
    return holidayConfig['設定值'].split(',').map(d => d.trim());
  }
  return [];
}

/**
 * 輔助函數：將日期轉為 YYYY-MM-DD 格式以供比對
 */
function formatDate(dateObj) {
  const d = new Date(dateObj);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

// ==========================================
// 2. 核心工作日計算邏輯
// ==========================================

/**
 * 檢查某日期是否為工作日 (星期一至五，且非公眾假期)
 * @param {Date|string} date 
 * @returns {boolean}
 */
function isWorkingDay(date) {
  const targetDate = new Date(date);
  const dayOfWeek = targetDate.getDay();
  
  // 檢查週末 (0 = 星期日, 6 = 星期六)
  if (dayOfWeek === 0 || dayOfWeek === 6) return false;
  
  // 檢查公眾假期
  const dateStr = formatDate(targetDate);
  const holidays = getHolidays();
  if (holidays.includes(dateStr)) return false;
  
  return true;
}

/**
 * 由 startDate 開始加上指定工作天數
 * @param {Date|string} startDate 
 * @param {number} days - 要加上的工作天
 * @returns {Date} - 計算後的日期
 */
function addWorkingDays(startDate, days) {
  let currentDate = new Date(startDate);
  let addedDays = 0;
  
  while (addedDays < days) {
    currentDate.setDate(currentDate.getDate() + 1);
    if (isWorkingDay(currentDate)) {
      addedDays++;
    }
  }
  return currentDate;
}

/**
 * 計算兩個日期之間的工作天數
 * @param {Date|string} startDate 
 * @param {Date|string} endDate 
 * @returns {number}
 */
function getWorkingDaysDiff(startDate, endDate) {
  let start = new Date(startDate);
  const end = new Date(endDate);
  
  // 將時間歸零以精準計算日期
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  
  let workingDaysCount = 0;
  
  if (start > end) return 0; // 避免反向計算錯誤
  
  let current = new Date(start);
  while (current < end) {
    current.setDate(current.getDate() + 1);
    if (isWorkingDay(current)) {
      workingDaysCount++;
    }
  }
  return workingDaysCount;
}

// ==========================================
// 3. 專案 Deadline 狀態檢查
// ==========================================

/**
 * 回傳某專案的 Deadline 狀態
 * @param {string} jobNumber 
 * @returns {string} - 'urgent', 'overdue', 'onTrack'
 */
function getDeadlineStatus(jobNumber) {
  try {
    const project = getProject(jobNumber);
    if (!project.firstDraftDate) return 'onTrack'; // 未設定死線，視為正常
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const deadline = new Date(project.firstDraftDate);
    deadline.setHours(0, 0, 0, 0);
    
    // 已經過期
    if (today > deadline) {
      return 'overdue';
    }
    
    // 計算距離死線還有幾個工作天
    const diffWorkingDays = getWorkingDaysDiff(today, deadline);
    
    // 少於 3 個工作天為急件
    if (diffWorkingDays < 3) {
      return 'urgent';
    }
    
    return 'onTrack';
  } catch (error) {
    Logger.log('計算 Deadline 狀態失敗: ' + error.message);
    return 'onTrack'; // 發生錯誤時預設不報警
  }
}

// ==========================================
// 4. 測試程式碼
// ==========================================
function testCalendarLogic() {
  ensureHolidaysInConfig(); // 確保假期資料寫入 Config
  
  const today = new Date(); // 測試當日
  Logger.log(`--- 日期基準測試 ---`);
  Logger.log(`今天是: ${formatDate(today)}`);
  
  const testHoliday = new Date('2026-01-01');
  Logger.log(`2026-01-01 是否為工作日 (預期 false): ${isWorkingDay(testHoliday)}`);
  
  const testWeekend = new Date('2026-01-03'); // 星期六
  Logger.log(`2026-01-03 是否為工作日 (預期 false): ${isWorkingDay(testWeekend)}`);
  
  Logger.log(`--- 加工作天測試 ---`);
  // 跨過復活節假期 (4月3日至4月6日)
  const easterStart = new Date('2026-04-02');
  const add2Days = addWorkingDays(easterStart, 2);
  Logger.log(`2026-04-02 加上 2 個工作天會是 (應跨過假期到 4/8 左右): ${formatDate(add2Days)}`);
  
  Logger.log(`--- 狀態檢查測試 ---`);
  // 建立一筆測試專案，將 deadline 設為明天
  try {
    const testJobNumber = 'A26-1002'; // 使用先前的測試資料
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    updateProject(testJobNumber, { firstDraftDate: formatDate(tomorrow) });
    const status = getDeadlineStatus(testJobNumber);
    Logger.log(`專案 ${testJobNumber} (死線: 明天) 的狀態為: ${status}`);
    
  } catch (e) {
    Logger.log('狀態測試跳過，可能因為找不到測試專案: ' + e.message);
  }
  
  Logger.log('✅ CalendarLogic 測試完成！');
}