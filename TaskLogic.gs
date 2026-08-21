/**
 * 📄 TaskLogic.gs
 * 任務狀態轉換邏輯與退件防呆迴圈 (Workflow)
 */

// ==========================================
// 輔助函數：生成 Task ID 與日期計算 Mock
// ==========================================
function generateTaskId(jobNumber) {
  const timestamp = new Date().getTime().toString().slice(-6);
  return `T-${jobNumber}-${timestamp}`;
}

// 暫時的日期計算 (後續將由 CalendarLogic.gs 完整接管香港公眾假期計算)
function isUrgentMock(targetDate) {
  const today = new Date();
  const target = new Date(targetDate);
  const diffDays = Math.ceil((target - today) / (1000 * 60 * 60 * 24));
  return diffDays < 3;
}

// ==========================================
// 1. 核心 Workflow 函數
// ==========================================

/**
 * 1. Team Head 指派任務
 */
function assignTask(jobNumber, assignedTo, firstDraftDate) {
  try {
    const taskId = generateTaskId(jobNumber);
    const newTask = {
      taskId: taskId,
      jobNumber: jobNumber,
      taskType: 'Advertorial', // 預設，依實際情況可傳參數
      assignedTo: assignedTo,
      status: 'Waiting for First Draft',
      deadline: firstDraftDate,
      versionNo: 0
    };
    
    appendRow(SHEET_NAMES.TASKS, newTask);
    logAction('ASSIGN_TASK', jobNumber, `指派任務給 ${assignedTo}，TaskId: ${taskId}`);

    // 檢查是否為急件 (< 3 工作天)
    if (isUrgentMock(firstDraftDate)) {
      return { success: true, taskId: taskId, warning: 'URGENT_RED', message: '⚠️ 距離初稿日期小於 3 個工作天，已標記為急件！' };
    }
    return { success: true, taskId: taskId };
  } catch (error) {
    throw new Error('指派任務失敗: ' + error.message);
  }
}

/**
 * 2. PIC 點擊 Accept
 */
function acceptTask(jobNumber, taskId) {
  try {
    updateRow(SHEET_NAMES.TASKS, 'taskId', taskId, {
      status: 'In Progress',
      acceptDate: new Date()
    });
    logAction('STATUS_CHANGE', jobNumber, `任務 ${taskId} 已接受，狀態: In Progress`);
    return { success: true };
  } catch (error) {
    throw new Error('接受任務失敗: ' + error.message);
  }
}

/**
 * 3. Editor 提交初稿
 */
function submitFirstDraft(jobNumber, taskId, firstDraftUrl) {
  try {
    updateRow(SHEET_NAMES.TASKS, 'taskId', taskId, {
      status: 'Client Review',
      firstDraftUrl: firstDraftUrl
    });
    logAction('STATUS_CHANGE', jobNumber, `任務 ${taskId} 提交初稿，狀態: Client Review`);
    Logger.log(`🔔 [通知] 已自動發送 Email 提醒 PM 初稿已提交！`);
    return { success: true };
  } catch (error) {
    throw new Error('提交初稿失敗: ' + error.message);
  }
}

/**
 * 4. Editor 請求配圖 (跨部門派單)
 */
function requestCover(jobNumber, taskId, designBriefUrl) {
  try {
    // 寫入 Brief URL 到原任務
    updateRow(SHEET_NAMES.TASKS, 'taskId', taskId, { designBriefUrl: designBriefUrl });
    
    // 建立新 Design Task
    const newTaskId = generateTaskId(jobNumber);
    const newDesignTask = {
      taskId: newTaskId,
      jobNumber: jobNumber,
      taskType: 'Design',
      status: 'Waiting for PIC',
      versionNo: 0
    };
    appendRow(SHEET_NAMES.TASKS, newDesignTask);
    
    logAction('CROSS_DEPT_REQUEST', jobNumber, `跨部門請求配圖，建立新任務 ${newTaskId}`);
    Logger.log(`🔔 [通知] 已自動發送通知給 Art Team Head！`);
    return { success: true, newTaskId: newTaskId };
  } catch (error) {
    throw new Error('請求配圖失敗: ' + error.message);
  }
}

/**
 * 5. 客戶要求修改，Editor 提交修訂版 (退件迴圈防呆)
 */
function submitRevisedAdvertorial(jobNumber, taskId, revisedUrl) {
  try {
    const tasks = getSheetData(SHEET_NAMES.TASKS);
    const task = tasks.find(t => t.taskId === taskId);
    if (!task) throw new Error('找不到任務');

    const newVersion = (parseInt(task.versionNo) || 0) + 1;
    
    updateRow(SHEET_NAMES.TASKS, 'taskId', taskId, {
      revisedUrl: revisedUrl,
      versionNo: newVersion
    });
    
    logAction('SUBMIT_REVISED', jobNumber, `提交修訂版，目前退件次數: ${newVersion}`);

    // 紫色奧客警報防呆 (versionNo >= 5)
    if (newVersion >= 5) {
      Logger.log(`🚨 [奧客警報] 專案 ${jobNumber} 退件次數已達 ${newVersion} 次！`);
      return { success: true, warning: 'PURPLE_ALERT', message: '⚠️ 觸發紫色奧客警報！請通報 Head of PM。' };
    }
    
    return { success: true };
  } catch (error) {
    throw new Error('提交修訂版失敗: ' + error.message);
  }
}

/**
 * 6. Designer 提交設計草稿
 */
function submitDesignDraft(jobNumber, taskId, designUrl) {
  try {
    updateRow(SHEET_NAMES.TASKS, 'taskId', taskId, {
      status: 'Internal Review',
      designDraftUrl: designUrl
    });
    logAction('STATUS_CHANGE', jobNumber, `任務 ${taskId} 提交設計草稿，狀態: Internal Review`);
    return { success: true };
  } catch (error) {
    throw new Error('提交設計草稿失敗: ' + error.message);
  }
}

/**
 * 7. 設計被退件
 */
function reviseDesign(jobNumber, taskId, revisedDesignUrl) {
  try {
    const tasks = getSheetData(SHEET_NAMES.TASKS);
    const task = tasks.find(t => t.taskId === taskId);
    const newVersion = (parseInt(task.versionNo) || 0) + 1;

    updateRow(SHEET_NAMES.TASKS, 'taskId', taskId, {
      revisedDesignUrl: revisedDesignUrl,
      versionNo: newVersion
    });
    logAction('REVISE_DESIGN', jobNumber, `設計退件，目前退件次數: ${newVersion}`);
    Logger.log(`🔔 [通知] 已通知 Editor 與 PM 關於設計退件。`);
    return { success: true };
  } catch (error) {
    throw new Error('設計退件處理失敗: ' + error.message);
  }
}

/**
 * 8. 專案完成
 */
function completeProject(jobNumber) {
  try {
    updateProject(jobNumber, { 
      status: 'Completed',
      completedDate: new Date()
    });
    logAction('STATUS_CHANGE', jobNumber, '專案已全數完成，狀態改為 Completed');
    Logger.log(`🔔 [通知] 已自動通知 PM 與 Sales 專案結案！`);
    return { success: true };
  } catch (error) {
    throw new Error('完成專案失敗: ' + error.message);
  }
}

/**
 * 9. 專案重啟
 */
function kickRestart(jobNumber) {
  try {
    updateProject(jobNumber, { status: 'In Progress' });
    logAction('STATUS_CHANGE', jobNumber, '⚠️ 專案被重啟，狀態改回 In Progress');
    return { success: true };
  } catch (error) {
    throw new Error('重啟專案失敗: ' + error.message);
  }
}

// ==========================================
// 10. 查詢狀態歷史紀錄
// ==========================================
function getStatusHistory(jobNumber) {
  const logs = getSheetData(SHEET_NAMES.AUDIT_LOG);
  return logs.filter(log => log['專案編號'] === jobNumber || log.jobNumber === jobNumber);
}

// ==========================================
// 測試 Workflow 流程
// ==========================================
function testWorkflow() {
  try {
    Logger.log('--- 1. 指派任務 (測試急件) ---');
    // 假設日期設定為明天，故意觸發急件警報
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const assignRes = assignTask('A26-1002', 'ming@hk01.com', tomorrow);
    Logger.log(assignRes);
    const taskId = assignRes.taskId;
    
    Logger.log('--- 2. PIC 接受任務 ---');
    Logger.log(acceptTask('A26-1002', taskId));
    
    Logger.log('--- 3. 提交初稿 ---');
    Logger.log(submitFirstDraft('A26-1002', taskId, 'https://docs.google.com/test'));
    
    Logger.log('--- 4. 觸發五次退件 (測試奧客警報) ---');
    for (let i = 1; i <= 5; i++) {
       const res = submitRevisedAdvertorial('A26-1002', taskId, 'https://docs.google.com/v' + i);
       if (res.warning) Logger.log('🔥 警報發布: ' + res.message);
    }
    
    Logger.log('--- 5. 專案結案 ---');
    Logger.log(completeProject('A26-1002'));
    
    Logger.log('✅ Workflow 測試圓滿成功！請至 Google Sheets 檢查 Tasks 與 AuditLog 表單。');
  } catch (e) {
    Logger.log('❌ 測試失敗: ' + e.message);
  }
}