// ==========================================
// 2026 年 Active & Pending 專案極速批次匯入工具
// ==========================================

function import2026ActivePendingJobs() {
  const SPREADSHEET_NAME = 'Job Status Report_PMO'; 
  
  try {
    const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
    if (!files.hasNext()) {
      throw new Error(`在 Google Drive 中找不到名為 "${SPREADSHEET_NAME}" 的檔案！請確認檔案已成功更名並儲存在 Drive 中。`);
    }

    const file = files.next();
    const sourceSS = SpreadsheetApp.open(file);
    const sourceSheet = sourceSS.getSheets()[0];
    const sourceData = sourceSheet.getDataRange().getValues();

    const targetSS = SpreadsheetApp.getActiveSpreadsheet();
    const projectsSheet = targetSS.getSheetByName(SHEET_NAMES.PROJECTS);
    const tasksSheet = targetSS.getSheetByName(SHEET_NAMES.TASKS);

    const projectsBatch = [];
    const tasksBatch = [];

    // 從第 3 行開始 (index 2)
    for (let i = 2; i < sourceData.length; i++) {
      const row = sourceData[i];
      
      const subDate = row[0];
      const jobNature = row[1];
      const jobStatus = String(row[2] || '').trim();
      const soStatus = row[3];
      const jobNumber = String(row[4] || '').trim();
      const client = row[5];
      const product = row[6];
      const launchDate = row[7];
      const textJobType = row[8];
      const itemType = row[9];
      const sales = row[10];
      const pm = row[11];
      const editor = row[13];

      // 🎯 嚴格過濾：2026 年 且 Status 僅限 Active 或 Pending
      const is2026 = jobNumber.toUpperCase().includes('A26') || 
                     (subDate instanceof Date && subDate.getFullYear() === 2026);
      const isActiveOrPending = ['Active', 'Pending'].includes(jobStatus);

      if (is2026 && isActiveOrPending) {
        // 1. 收集 Project 資料
        projectsBatch.push([
          jobNumber,
          subDate ? new Date(subDate) : new Date(),
          launchDate || 'TBC',
          jobNature || 'Confirmed',
          jobStatus,
          soStatus || '',
          client || 'Unknown Client',
          product || 'General Campaign',
          textJobType || 'Advertorial',
          itemType || '',
          sales || pm || 'Sales',
          pm || 'PM'
        ]);

        // 2. 收集 Task 資料 (若有指定 Editor)
        if (editor) {
          tasksBatch.push([
            `T-${jobNumber}-${projectsBatch.length}`,
            jobNumber,
            textJobType || 'Advertorial',
            editor,
            jobStatus === 'Active' ? 'In Progress' : 'Pending',
            launchDate instanceof Date ? new Date(launchDate) : new Date(),
            '', '', '', '', '', ''
          ]);
        }
      }
    }

    // 🚀 整批寫入 (Batch Insert)，約 2~3 秒內完成！
    if (projectsBatch.length > 0) {
      projectsSheet.getRange(projectsSheet.getLastRow() + 1, 1, projectsBatch.length, projectsBatch[0].length).setValues(projectsBatch);
    }
    if (tasksBatch.length > 0) {
      tasksSheet.getRange(tasksSheet.getLastRow() + 1, 1, tasksBatch.length, tasksBatch[0].length).setValues(tasksBatch);
    }

    const summaryMsg = `🎉 極速匯入成功！\n• 專案數 (Projects): ${projectsBatch.length} 筆\n• 任務數 (Tasks): ${tasksBatch.length} 筆`;
    SpreadsheetApp.getUi().alert('匯入完成', summaryMsg, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getUi().alert('匯入失敗', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}