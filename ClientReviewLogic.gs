/**
 * 📄 ClientReviewLogic.gs
 * Client Review 追蹤與客戶催促 (Reminder) 功能
 */

// ==========================================
// 1. 計算等待日數
// ==========================================
/**
 * 計算某專案已等待客戶回覆多少日
 * @param {string} jobNumber 
 * @returns {number} 等待日數 (日曆日)
 */
function getWaitingDays(jobNumber) {
  try {
    const logs = getSheetData(SHEET_NAMES.AUDIT_LOG);
    
    // 找出該專案的所有 log
    // 支援中英文欄位名稱，相容 AuditLog 的表頭設定
    const projectLogs = logs.filter(log => log['專案編號'] === jobNumber || log.jobNumber === jobNumber);
    
    let reviewStartDate = new Date(); // 預設今天
    let found = false;

    // 從最新的 Log 往回找，尋找最後一次進入 'Client Review' 的時間
    for (let i = projectLogs.length - 1; i >= 0; i--) {
      const log = projectLogs[i];
      const details = String(log['詳細內容'] || log.details || '');
      
      // 假設在 TaskLogic 中，狀態變更為 Client Review 會寫入 Log
      if (details.includes('Client Review')) {
        // 取得時間戳記 (相容不同欄位名稱)
        const timestamp = log['時間'] || log.timestamp || log.Date || Object.values(log)[0];
        reviewStartDate = new Date(timestamp);
        found = true;
        break;
      }
    }
    
    // 如果在 Log 中找不到，退而求其次使用 Project 表中的 updatedAt
    if (!found) {
      const project = getProject(jobNumber);
      if (project && project.updatedAt) {
        reviewStartDate = new Date(project.updatedAt);
      }
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    reviewStartDate.setHours(0, 0, 0, 0);
    
    // 計算相差日數 (日曆日，因為客戶假日也可能被計算在等待時間內)
    const diffTime = Math.abs(today - reviewStartDate);
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    
    return diffDays;
  } catch (error) {
    Logger.log(`計算 ${jobNumber} 等待日數失敗: ` + error.message);
    return 0;
  }
}

// ==========================================
// 2. 獲取 Client Review 清單
// ==========================================
/**
 * 列出所有處於 Client Review 狀態的專案
 * @returns {Object} JSON 格式的清單
 */
function getClientReviewList() {
  try {
    const tasks = getSheetData(SHEET_NAMES.TASKS);
    const projects = getSheetData(SHEET_NAMES.PROJECTS);
    
    // 找出所有狀態為 'Client Review' 的任務
    const reviewTasks = tasks.filter(t => t.status === 'Client Review');
    
    // 取得唯一的 jobNumber (避免同一個專案有多個任務在 review)
    const uniqueJobNumbers = [...new Set(reviewTasks.map(t => t.jobNumber))];
    
    const reviewList = uniqueJobNumbers.map(jobNumber => {
      const proj = projects.find(p => p.jobNumber === jobNumber) || {};
      const waitingDays = getWaitingDays(jobNumber);
      
      return {
        jobNumber: jobNumber,
        client: proj.clientName || '未知客戶',
        product: proj.productName || '未填寫',
        salesPerson: proj.salesPerson || '未指派',
        pmName: proj.pmName || '未指派',
        waitingDays: waitingDays
      };
    });
    
    // 按等待日數由大到小排序 (最耐嘅排最前)
    reviewList.sort((a, b) => b.waitingDays - a.waitingDays);
    
    return { success: true, data: reviewList };
  } catch (error) {
    return { success: false, message: '獲取 Client Review 清單失敗: ' + error.message };
  }
}

// ==========================================
// 3. 一鍵發送催促通知
// ==========================================
/**
 * 向負責的 Sales 或 PM 發送催促文字，讓他們轉發給客戶
 * @param {string} jobNumber 
 * @param {string} method - 'email' 或 'whatsapp'
 */
function sendClientReminder(jobNumber, method) {
  try {
    const project = getProject(jobNumber);
    if (!project) throw new Error('找不到專案資料');

    // 優先找負責對外的 Sales，沒有的話找 PM
    const targetName = project.salesPerson || project.pmName;
    if (!targetName) throw new Error('專案未設定 Sales 或 PM，無法通知');

    const user = getUserByName(targetName);
    if (!user) throw new Error(`系統找不到同事「${targetName}」的聯絡資料`);

    // 準備給客戶的預設催促訊息
    const clientMessage = `親愛的客戶，\n\n關於專案 ${jobNumber}（${project.productName}）目前正等待您的回覆，請盡快確認以利後續進度。\n\n謝謝！`;

    if (method === 'email') {
      const subject = `[催促轉發] 專案 ${jobNumber} 等待客戶確認中`;
      const htmlBody = `
        <p>Hi ${targetName},</p>
        <p>系統偵測到專案 <b>${jobNumber}</b> 卡在 Client Review 階段。請複製以下文字轉發給客戶催促進度：</p>
        <div style="background-color:#f9f9f9; padding: 15px; border-left: 4px solid #f2994a; margin-top:10px;">
          <pre style="font-family: inherit; margin:0; white-space: pre-wrap;">${clientMessage}</pre>
        </div>
      `;
      sendEmailNotification(user.email, subject, htmlBody);
      
    } else if (method === 'whatsapp') {
      if (!user.phoneNumber) throw new Error(`同事 ${targetName} 未在 Users 表中設定電話號碼`);
      const waMsg = `🔔 [請轉發客戶]\n專案 ${jobNumber} 等待確認中。請複製以下內容發給客戶：\n\n---\n${clientMessage}`;
      sendWhatsAppMessage(user.phoneNumber, waMsg);
      
    } else {
      throw new Error('不支援的發送方式 (僅支援 email 或 whatsapp)');
    }

    // 記錄 AuditLog
    logAction('CLIENT_REMINDER', jobNumber, `使用 ${method.toUpperCase()} 發送催促文字給 ${targetName}`);
    
    return { success: true, message: `已成功將催促訊息發送給 ${targetName}` };
  } catch (error) {
    throw new Error('發送催促失敗: ' + error.message);
  }
}

// ==========================================
// 4. 測試程式碼
// ==========================================
function testClientReviewLogic() {
  try {
    Logger.log('--- 1. 為了測試，模擬一個 Client Review 的任務 ---');
    // 假設我們在 Tasks 表裡插一個狀態是 Client Review 的測試任務
    const testTaskId = 'T-TEST-' + new Date().getTime();
    appendRow(SHEET_NAMES.TASKS, {
      taskId: testTaskId,
      jobNumber: 'A26-1002', // 用之前的測試專案
      status: 'Client Review',
      assignedTo: Session.getActiveUser().getEmail()
    });
    // 寫入一筆 Log 製造等待日數
    logAction('STATUS_CHANGE', 'A26-1002', `任務 ${testTaskId} 提交初稿，狀態: Client Review`);
    
    Logger.log('--- 2. 測試獲取清單 ---');
    const reviewList = getClientReviewList();
    Logger.log(JSON.stringify(reviewList, null, 2));
    
    Logger.log('--- 3. 測試一鍵催促 (Email) ---');
    if (reviewList.data && reviewList.data.length > 0) {
      // 在這個測試中，A26-1002 的 PM 是 Wincy，如果 Users 找不到 Wincy 會報錯，這是預期內的防呆機制
      // 為了測試順利，您可以確認 Users 表裡有 Wincy 或是把專案的 PM 改成自己
      const result = sendClientReminder('A26-1002', 'email');
      Logger.log(result.message);
    }
    
    Logger.log('✅ ClientReviewLogic 測試完成！');
  } catch (e) {
    Logger.log('❌ 測試失敗: ' + e.message);
  }
}

// ==========================================
// 💡 全新：獲取客戶審批 (Client Review) 追蹤資料 API (智能搜尋 In Progress 關卡版)
// ==========================================
function api_getClientReviewData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if(!sheet) return { success: true, data: [] };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxSales = headers.findIndex(h => h.includes('sales'));
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    
    let results = [];
    
    for(let i=1; i < data.length; i++) {
      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : '';
      if(pStatus === 'Completed' || pStatus === 'Recycle Bin' || pStatus === 'Cancelled') continue;
      
      let wfData = {};
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables')) { try { wfData = JSON.parse(cellStr); break; } catch(e){} }
      }
      
      if(wfData && wfData.deliverables) {
        wfData.deliverables.forEach(d => {
          if(d.status === 'Completed' || d.status === 'Deleted') return;
          
          // 💡 核心修復：不要依賴可能卡住的 currentStep，直接抓出狀態為 'In Progress' 的那一關！
          const currentStepObj = d.workflow ? d.workflow.find(s => s.status === 'In Progress') : null;
          
          if (currentStepObj && (currentStepObj.reviewStatus === 'Reviewing' || currentStepObj.name.toLowerCase().includes('client') || currentStepObj.dept.toLowerCase().includes('client'))) {
             
             let waitDays = 1;
             try {
                 // 調用你原本寫好的 getWaitingDays 來計算天數
                 waitDays = getWaitingDays(data[i][idxJobNum]);
             } catch(e) {}

             results.push({
               jobNumber: idxJobNum >= 0 ? data[i][idxJobNum] : '',
               client: idxClient >= 0 ? data[i][idxClient] : '',
               taskName: d.name || '未命名項目',
               stepName: currentStepObj.name || ('Step ' + currentStepObj.step),
               pmName: idxPM >= 0 ? data[i][idxPM] : '未指定',
               salesName: idxSales >= 0 ? data[i][idxSales] : '未指定',
               waitingDays: waitDays
             });
          }
        });
      }
    }
    
    // 按等待日數由大到小排序
    results.sort((a, b) => b.waitingDays - a.waitingDays);
    
    return { success: true, data: results };
  } catch(e) {
    return { success: false, message: e.message };
  }
}