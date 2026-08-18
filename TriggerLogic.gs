/**
 * 📄 TriggerLogic.gs
 * Time Trigger 排程任務 (Cron Jobs)
 * 包含每日檢查、報表派發與自動清理
 */

// ==========================================
// 1. 每日檢查逾期任務 (Daily Check)
// ==========================================
function checkOverdueTasks() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TASKS);
    const tasks = getSheetData(SHEET_NAMES.TASKS);
    const projects = getSheetData(SHEET_NAMES.PROJECTS);
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let blockedCount = 0;
    let notifyCount = 0;

    // 從第 2 列開始 (第 1 列是標題)
    for (let i = 0; i < tasks.length; i++) {
      const t = tasks[i];
      if (t.status === 'Completed' || t.status === 'Blocked') continue; // 已經完成或卡關的跳過

      const proj = projects.find(p => p.jobNumber === t.jobNumber) || {};
      if (!proj.firstDraftDate) continue;

      const draftDate = new Date(proj.firstDraftDate);
      draftDate.setHours(0, 0, 0, 0);

      // 判斷是否逾期
      if (today > draftDate) {
        const overdueDays = getWorkingDaysDiff(draftDate, today);
        const pmEmail = getUserByName(proj.pmName)?.email;
        const picEmail = t.assignedTo;
        
        // 條件 A：逾期超過 3 個工作天 -> 強制轉為 Blocked
        if (overdueDays > 3) {
          // 找出 'status' 在表中的欄位索引 (A=1, B=2...)
          const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const statusColIndex = headers.indexOf('status') + 1;
          
          // 更新 Sheet 中的狀態 (i + 2 因為陣列是 0-based，且要跳過標題列)
          sheet.getRange(i + 2, statusColIndex).setValue('Blocked');
          
          logAction('AUTO_UPDATE', t.jobNumber, `任務 ${t.taskId} 逾期超過 3 日，系統自動標記為 Blocked`);
          notifyStatusChange(t.jobNumber, t.status, 'Blocked'); // 觸發通知
          blockedCount++;
        } 
        // 條件 B：逾期 1~3 天內 -> 發送警告 Email
        else {
          const subject = `⚠️ [逾期警告] 專案 ${t.jobNumber} 已逾期 ${overdueDays} 個工作天`;
          const body = `<p>專案 <b>${t.jobNumber} (${proj.clientName})</b> 已超過初稿死線 (${formatDate(draftDate)})。</p><p>目前負責人：${t.assignedTo}</p><p>請盡速處理，若逾期超過 3 日系統將自動凍結此任務 (Blocked)。</p>`;
          
          if (pmEmail) sendEmailNotification(pmEmail, subject, body);
          if (picEmail && picEmail !== pmEmail) sendEmailNotification(picEmail, subject, body);
          notifyCount++;
        }
      }
    }
    Logger.log(`每日逾期檢查完成：凍結 ${blockedCount} 件，警告發送 ${notifyCount} 件。`);
  } catch (error) {
    Logger.log('執行 checkOverdueTasks 發生錯誤: ' + error.message);
  }
}

// ==========================================
// 2. 每週報表生成與發送 (Weekly Report)
// ==========================================
function generateWeeklyReport() {
  try {
    const today = new Date();
    // 計算上週的週一到週日
    const lastWeekEnd = new Date(today.getTime() - (today.getDay() || 7) * 24 * 60 * 60 * 1000);
    const lastWeekStart = new Date(lastWeekEnd.getTime() - 6 * 24 * 60 * 60 * 1000);

    const startDateStr = formatDate(lastWeekStart);
    const endDateStr = formatDate(lastWeekEnd);

    // 獲取數據
    const lateData = calculateLateRate(startDateStr, endDateStr);
    const deptData = getDepartmentPerformance();

    // 組裝 HTML 報表
    let htmlBody = `<h2 style="color: #005088;">週報：PMO 系統績效分析</h2>`;
    htmlBody += `<p>統計期間：${startDateStr} ~ ${endDateStr}</p>`;
    
    htmlBody += `<h3>1. 全公司完件逾期率</h3>`;
    if (lateData.success) {
      htmlBody += `<p>整體逾期率：<b>${lateData.data.overallLateRate}%</b></p><ul>`;
      lateData.data.pmBreakdown.forEach(pm => {
        htmlBody += `<li>${pm.pmName}：${pm.lateRate}% (${pm.lateCount}/${pm.totalCount} 件)</li>`;
      });
      htmlBody += `</ul>`;
    }

    htmlBody += `<h3>2. 各部門當前負載</h3><table border="1" cellpadding="5" style="border-collapse: collapse;">`;
    htmlBody += `<tr style="background-color: #f2f2f2;"><th>部門</th><th>活躍任務</th><th>平均退件數</th></tr>`;
    if (deptData.success) {
      deptData.data.forEach(d => {
        htmlBody += `<tr><td>${d.department}</td><td>${d.totalTasks}</td><td>${d.avgRevisionCount}</td></tr>`;
      });
    }
    htmlBody += `</table><br><p>詳細數據請登入 PMO Web App 檢視 Management View。</p>`;

    // 找出所有 Management 角色並發送
    const users = getSheetData(SHEET_NAMES.USERS);
    const mgmtUsers = users.filter(u => u.role === 'Management');
    
    mgmtUsers.forEach(u => {
      sendEmailNotification(u.email, `[HK01 PMO] 系統每週績效報表 (${startDateStr})`, htmlBody);
    });

    logAction('WEEKLY_REPORT', 'ALL', `已生成並發送週報給 ${mgmtUsers.length} 位管理層`);
  } catch (error) {
    Logger.log('執行 generateWeeklyReport 發生錯誤: ' + error.message);
  }
}

// ==========================================
// 3. 清理舊日誌 (Cleanup Old Logs)
// ==========================================
function cleanupOldLogs() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.AUDIT_LOG);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers = data[0];
    const timeIndex = headers.indexOf('時間') !== -1 ? headers.indexOf('時間') : headers.indexOf('timestamp');
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 90); // 90 天前

    const newData = [headers];
    for (let i = 1; i < data.length; i++) {
      const rowTime = new Date(data[i][timeIndex]);
      if (rowTime >= cutoffDate) {
        newData.push(data[i]);
      }
    }

    // 只有在真的有舊資料被過濾掉時才執行寫入 (節省系統資源)
    if (newData.length < data.length) {
      sheet.clearContents();
      sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      logAction('SYSTEM_CLEANUP', 'ALL', `已清理 ${data.length - newData.length} 筆超過 90 天的舊日誌`);
    } else {
      Logger.log('沒有超過 90 天的日誌需要清理。');
    }
  } catch (error) {
    Logger.log('執行 cleanupOldLogs 發生錯誤: ' + error.message);
  }
}

// ==========================================
// 4. Trigger 安裝與解除工具 (Setup & Delete)
// ==========================================
/**
 * 一鍵安裝所有排程任務
 * 警告：執行前請先確認您用的是專屬系統帳號 (如 pmo-system@hk01.com)
 */
function setupTriggers() {
  deleteAllTriggers(); // 先刪除舊的避免重複
  
  // 1. 每日早上 8~9 點：檢查逾期任務 (在 Daily Digest 之前執行，這樣 Digest 就能抓到最新的 Blocked 狀態)
  ScriptApp.newTrigger('checkOverdueTasks')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  // 2. 每日早上 9~10 點：發送每日任務推送 (呼叫 NotificationLogic 中的函數)
  ScriptApp.newTrigger('sendDailyDigest')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // 3. 每週一早上 8~9 點：發送管理層週報
  ScriptApp.newTrigger('generateWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  // 4. 每週日凌晨 2~3 點：清理 90 天前舊日誌
  ScriptApp.newTrigger('cleanupOldLogs')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();

  Logger.log('✅ 所有自動化排程 (Triggers) 已成功建立！');
}

/**
 * 刪除當前帳號設定的所有 Triggers
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  Logger.log('🗑️ 已清除所有舊的 Triggers。');
}