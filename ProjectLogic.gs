/**
 * 📄 ProjectLogic.gs
 * 專案核心邏輯與自動編號生成
 */

// ==========================================
// 1. 輔助函數：解析 Job Number 年份
// ==========================================
/**
 * 解析現有 Job Number 的年份 (例如 'A26-1001' 會回傳 '26')
 * @param {string} jobNumber - 現有的專案編號
 * @returns {string|null} - 回傳兩位數年份，若格式錯誤則回傳 null
 */
function getYearFromJobNumber(jobNumber) {
  if (!jobNumber || typeof jobNumber !== 'string') return null;
  
  // 使用正則表達式匹配 A(兩位數字)-(四位數字)
  const match = jobNumber.match(/^A(\d{2})-\d{4,}$/);
  return match ? match[1] : null;
}

// ==========================================
// 2. 主函數：生成 Job Number
// ==========================================
/**
 * 自動生成下一個 Job Number (格式：A26-1001)
 * 規則：A + 當前西元年後兩碼 + 流水號 (從 1001 開始)
 * @returns {string} - 新的專案編號
 */
function generateJobNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects'); // 對應剛剛建立的 Projects 表
  if (!sheet) {
    throw new Error('找不到 Projects 表格，請確認是否已經執行 initializeSheets。');
  }

  // 取得當前年份後兩位 (例如 2026 -> "26")
  const currentYear = new Date().getFullYear().toString().slice(-2);
  const targetPrefix = `A${currentYear}-`;

  const data = sheet.getDataRange().getValues();
  let maxSequence = 1000; // 基準流水號 (代表今年第一筆會是 1001)

  // 遍歷 Projects 表格 (跳過第一行標題列)
  for (let i = 1; i < data.length; i++) {
    const existingJobNumber = data[i][0]; // jobNumber 欄位固定在第 A 欄 (Index 0)
    
    // 檢查是否為當年度的 Job Number
    if (existingJobNumber && typeof existingJobNumber === 'string' && existingJobNumber.startsWith(targetPrefix)) {
      // 切割字串取得流水號部分 (例如 'A26-1316' -> '1316')
      const sequenceStr = existingJobNumber.split('-')[1];
      const sequenceNum = parseInt(sequenceStr, 10);
      
      // 比對並找出當前最大值
      if (!isNaN(sequenceNum) && sequenceNum > maxSequence) {
        maxSequence = sequenceNum;
      }
    }
  }

  // 流水號 + 1，組成全新的 Job Number
  const nextSequence = maxSequence + 1;
  return `${targetPrefix}${nextSequence}`;
}

// ==========================================
// 3. 測試程式碼 (示範如何呼叫)
// ==========================================
/**
 * 測試生成與解析 Job Number
 */
function testGenerateJobNumber() {
  try {
    const newJobNumber = generateJobNumber();
    Logger.log('🎉 測試成功！為您生成的最新 Job Number 是: ' + newJobNumber);
    
    const parsedYear = getYearFromJobNumber(newJobNumber);
    Logger.log('🔍 解析函數測試：成功解析出年份 [' + parsedYear + ']');
  } catch (error) {
    Logger.log('❌ 發生錯誤: ' + error.message);
  }
}