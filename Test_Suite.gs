// ==========================================
// 系統健康自動化測試套件 (派發中心)
// ==========================================
function api_runSingleTest(testName) {
  try {
    if (testName === 'Test_PauseLog') {
      return runTest_PauseAndLog(); 
      
    } else if (testName === 'Test_RecycleBin') {
      return runTest_RecycleBin(); 
      
    } else if (testName === 'Test_AssignTask') {
      // 💡 就是這裡！把原本的「開發中」字眼刪掉，改成真正去呼叫底下的函數：
      return runTest_AssignTask(); 
      
    } else {
      throw new Error('未知的測試專案：' + testName);
    }
  } catch(e) {
    return { success: false, message: '❌ 測試派發中樞崩潰: ' + e.message };
  }
}

// ==========================================
// 測試模組 1：專案暫停與 Log 寫入驗證
// ==========================================
function runTest_PauseAndLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Projects');
  if (!sheet) return { success: false, message: '❌ 找不到 Projects 工作表' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
  const idxJobNum = headers.indexOf('jobNumber');
  const idxStatus = headers.indexOf('status');
  const idxTextJob = headers.indexOf('textJobType');

  if (idxJobNum === -1 || idxStatus === -1 || idxTextJob === -1) {
    return { success: false, message: '❌ 缺少 jobNumber, status 或 textJobType 欄位' };
  }

  // 1️⃣ 建立靶材
  const testJobNo = 'AUTO_TEST_' + new Date().getTime();
  let newRow = new Array(headers.length).fill('');
  newRow[idxJobNum] = testJobNo;
  newRow[idxStatus] = 'In Progress';
  newRow[idxTextJob] = ''; 
  sheet.appendRow(newRow);
  SpreadsheetApp.flush();

  let passed = true;
  let detailMsg = ['▶️ 開始執行 [測試 1：專案暫停與 Log]...'];

  try {
    // 2️⃣ 執行動作
    let payload = JSON.stringify({ jobNumber: testJobNo, action: 'PAUSE' });
    let res = api_manageProjectStatus(payload); 
    if (!res || !res.success) {
      detailMsg.push('  ❌ API 回傳失敗: ' + (res ? res.message : '無回應'));
      passed = false;
    }

    // 3️⃣ 驗證結果
    SpreadsheetApp.flush();
    let data = sheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][idxJobNum]) === testJobNo) {
        foundRow = i;
        break;
      }
    }

    if (foundRow !== -1) {
      let updatedStatus = data[foundRow][idxStatus];
      let updatedLog = data[foundRow][idxTextJob];

      if (updatedStatus === 'Paused') {
        detailMsg.push('  ✅ 狀態成功變更為 Paused');
      } else {
        detailMsg.push('  ❌ 狀態變更失敗');
        passed = false;
      }

      let logArray = [];
      try { logArray = JSON.parse(updatedLog); } catch(e){}
      if (Array.isArray(logArray) && logArray.length > 0 && logArray[0].action === 'Pause Project') {
        detailMsg.push('  ✅ Log 成功寫入 JSON 陣列');
      } else {
        detailMsg.push('  ❌ Log 寫入失敗 (發生靜默失敗)');
        passed = false;
      }
    } else {
      detailMsg.push('  ❌ 找不到測試專案，無法驗證');
      passed = false;
    }
  } catch (e) {
    detailMsg.push('  ❌ 執行過程崩潰: ' + e.message);
    passed = false;
  } finally {
    // 4️⃣ 清理靶材
    let finalData = sheet.getDataRange().getValues();
    for (let i = finalData.length - 1; i >= 0; i--) {
      if (String(finalData[i][idxJobNum]) === testJobNo) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    detailMsg.push('  🧹 測試靶材已銷毀');
  }

  let finalTitle = passed ? '🎉 【測試 1 通過】\n\n' : '⚠️ 【測試 1 失敗】\n\n';
  return { success: passed, message: finalTitle + detailMsg.join('\n') };
}

// ==========================================
// 測試模組 2：移至回收箱與還原測試
// ==========================================
function runTest_RecycleBin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Projects');
  if (!sheet) return { success: false, message: '❌ 找不到 Projects 工作表' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
  const idxJobNum = headers.indexOf('jobNumber');
  const idxStatus = headers.indexOf('status');
  const idxTextJob = headers.indexOf('textJobType');

  if (idxJobNum === -1 || idxStatus === -1 || idxTextJob === -1) {
    return { success: false, message: '❌ 缺少 jobNumber, status 或 textJobType 欄位' };
  }

  // 1️⃣ 建立靶材
  const testJobNo = 'AUTO_TEST_BIN_' + new Date().getTime();
  let newRow = new Array(headers.length).fill('');
  newRow[idxJobNum] = testJobNo;
  newRow[idxStatus] = 'In Progress';
  newRow[idxTextJob] = ''; 
  sheet.appendRow(newRow);
  SpreadsheetApp.flush();

  let passed = true;
  let detailMsg = ['▶️ 開始執行 [測試 2：移至回收箱與還原]...'];

  try {
    // 2️⃣ 執行動作 A：移至回收箱 (DELETE)
    let payloadDel = JSON.stringify({ jobNumber: testJobNo, action: 'DELETE' });
    let resDel = api_manageProjectStatus(payloadDel); 
    
    if (!resDel || !resDel.success) {
      detailMsg.push('  ❌ API (刪除) 回傳失敗: ' + (resDel ? resDel.message : '無回應'));
      passed = false;
    }

    // 3️⃣ 驗證結果 A：查水表確認是否進入 Recycle Bin
    SpreadsheetApp.flush();
    let data = sheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][idxJobNum]) === testJobNo) {
        foundRow = i;
        break;
      }
    }

    if (foundRow !== -1) {
      let updatedStatus = data[foundRow][idxStatus];
      let updatedLog = data[foundRow][idxTextJob];

      if (updatedStatus === 'Recycle Bin') {
        detailMsg.push('  ✅ [刪除階段] 狀態成功變更為 Recycle Bin');
      } else {
        detailMsg.push('  ❌ [刪除階段] 狀態變更失敗 (目前為 ' + updatedStatus + ')');
        passed = false;
      }

      let logArray = [];
      try { logArray = JSON.parse(updatedLog); } catch(e){}
      if (Array.isArray(logArray) && logArray.length > 0 && logArray[0].action === 'Delete Project') {
        detailMsg.push('  ✅ [刪除階段] Log 成功寫入 "Delete Project"');
      } else {
        detailMsg.push('  ❌ [刪除階段] Log 寫入失敗');
        passed = false;
      }
      
      // 4️⃣ 執行動作 B：從回收箱還原 (RESUME)
      if (passed) {
        let payloadRes = JSON.stringify({ jobNumber: testJobNo, action: 'RESUME' });
        api_manageProjectStatus(payloadRes); // 呼叫還原 API
        
        SpreadsheetApp.flush();
        let dataAfterResume = sheet.getDataRange().getValues();
        let statusAfterResume = dataAfterResume[foundRow][idxStatus];
        let logAfterResume = dataAfterResume[foundRow][idxTextJob];
        
        if (statusAfterResume === 'In Progress') {
          detailMsg.push('  ✅ [還原階段] 狀態成功恢復為 In Progress');
        } else {
          detailMsg.push('  ❌ [還原階段] 狀態恢復失敗 (目前為 ' + statusAfterResume + ')');
          passed = false;
        }
        
        let logArrayRes = [];
        try { logArrayRes = JSON.parse(logAfterResume); } catch(e){}
        if (Array.isArray(logArrayRes) && logArrayRes.length > 0 && logArrayRes[0].action === 'Resume Project') {
          detailMsg.push('  ✅ [還原階段] Log 成功寫入 "Resume Project"');
        } else {
          detailMsg.push('  ❌ [還原階段] Log 寫入失敗');
          passed = false;
        }
      }

    } else {
      detailMsg.push('  ❌ 找不到測試專案，無法驗證');
      passed = false;
    }
  } catch (e) {
    detailMsg.push('  ❌ 執行過程崩潰: ' + e.message);
    passed = false;
  } finally {
    // 5️⃣ 清理靶材
    let finalData = sheet.getDataRange().getValues();
    for (let i = finalData.length - 1; i >= 0; i--) {
      if (String(finalData[i][idxJobNum]) === testJobNo) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    detailMsg.push('  🧹 測試靶材已銷毀');
  }

  let finalTitle = passed ? '🎉 【測試 2 通過】\n\n' : '⚠️ 【測試 2 失敗】\n\n';
  return { success: passed, message: finalTitle + detailMsg.join('\n') };
}

// ==========================================
// 測試模組 3：指派任務與 Log 寫入驗證
// ==========================================
function runTest_AssignTask() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Projects');
  if (!sheet) return { success: false, message: '❌ 找不到 Projects 工作表' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
  const idxJobNum = headers.indexOf('jobNumber');
  const idxStatus = headers.indexOf('status');
  const idxTextJob = headers.indexOf('textJobType');

  if (idxJobNum === -1 || idxStatus === -1 || idxTextJob === -1) {
    return { success: false, message: '❌ 缺少 jobNumber, status 或 textJobType 欄位' };
  }

  // 1️⃣ 建立靶材：模擬一個剛開案的專案
  const testJobNo = 'AUTO_TEST_TASK_' + new Date().getTime();
  let newRow = new Array(headers.length).fill('');
  newRow[idxJobNum] = testJobNo;
  newRow[idxStatus] = 'In Progress';
  newRow[idxTextJob] = '[]'; // 建立一個空的 JSON 陣列
  sheet.appendRow(newRow);
  SpreadsheetApp.flush();

  let passed = true;
  let detailMsg = ['▶️ 開始執行 [測試 3：指派任務與 Log 寫入]...'];

  try {
  // 2️⃣ 執行動作：模擬指派任務 (呼叫真實的 API)
    // 根據 WorkflowUI.html，API 需要四個參數: jobNumber, deliverableId, stepNumber, assignee
    let resTask = api_assignStepAndStart(
      testJobNo,       // 測試專案編號
      'TEST_DELIV_1',  // 模擬的 Deliverable ID
      1,               // 模擬的 Step Number
      'Test_Assignee'  // 模擬的指派人員
    );

    
    if (!resTask || !resTask.success) {
      detailMsg.push('  ❌ API 回傳失敗: ' + (resTask ? resTask.message : '無回應'));
      detailMsg.push('  💡 提示：請在 Test_Suite.gs 的 runTest_AssignTask 中綁定正確的 API 函數！');
      passed = false;
    }

    // 3️⃣ 驗證結果：查水表確認 Log 寫入
    if (passed) {
      SpreadsheetApp.flush();
      let data = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][idxJobNum]) === testJobNo) {
          foundRow = i;
          break;
        }
      }

      if (foundRow !== -1) {
        let updatedLog = data[foundRow][idxTextJob];
        let logArray = [];
        try { logArray = JSON.parse(updatedLog); } catch(e){}
        
        // 驗證 Log 陣列中是否有 'Dispatch Task' 的紀錄
        let hasDispatchLog = logArray.some(log => log.action === 'Dispatch Task');

        if (hasDispatchLog) {
          detailMsg.push('  ✅ Log 成功寫入 "Dispatch Task" 紀錄');
        } else {
          detailMsg.push('  ❌ Log 寫入失敗 (未找到 Dispatch Task 紀錄)');
          passed = false;
        }
      } else {
        detailMsg.push('  ❌ 找不到測試專案，無法驗證');
        passed = false;
      }
    }
  } catch (e) {
    detailMsg.push('  ❌ 執行過程崩潰: ' + e.message);
    passed = false;
  } finally {
    // 4️⃣ 清理靶材
    let finalData = sheet.getDataRange().getValues();
    for (let i = finalData.length - 1; i >= 0; i--) {
      if (String(finalData[i][idxJobNum]) === testJobNo) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    detailMsg.push('  🧹 測試靶材已銷毀');
  }

  let finalTitle = passed ? '🎉 【測試 3 通過】\n\n' : '⚠️ 【測試 3 失敗】\n\n';
  return { success: passed, message: finalTitle + detailMsg.join('\n') };
}