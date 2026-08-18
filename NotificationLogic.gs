/**
 * 📄 NotificationLogic.gs
 * 通知中心 (Email + WhatsApp) 與自動推播排程
 */

// ==========================================
// 1. 輔助函數：讀取設定與使用者聯絡資料
// ==========================================

/**
 * 確保 WhatsApp API 設定在 Config 表中
 */
function ensureWhatsAppConfig() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  const configData = getSheetData(SHEET_NAMES.CONFIG);
  
  const tokenExists = configData.find(c => c['設定名稱'] === 'WHATSAPP_API_TOKEN' || c.key === 'WHATSAPP_API_TOKEN');
  const phoneIdExists = configData.find(c => c['設定名稱'] === 'WHATSAPP_PHONE_ID' || c.key === 'WHATSAPP_PHONE_ID');
  
  if (!tokenExists) configSheet.appendRow(['WHATSAPP_API_TOKEN', 'YOUR_TOKEN_HERE', 'Meta WhatsApp Cloud API Token']);
  if (!phoneIdExists) configSheet.appendRow(['WHATSAPP_PHONE_ID', 'YOUR_PHONE_ID_HERE', 'Meta WhatsApp Phone Number ID']);
}

/**
 * 透過姓名尋找使用者資料 (因為 Projects 表存的是姓名)
 */
function getUserByName(name) {
  const users = getSheetData(SHEET_NAMES.USERS);
  return users.find(u => u.name === name);
}

/**
 * 透過 Email 尋找使用者資料
 */
function getUserByEmail(email) {
  const users = getSheetData(SHEET_NAMES.USERS);
  return users.find(u => u.email === email);
}

// ==========================================
// 2. 基礎發送函數 (Base Senders)
// ==========================================

/**
 * 發送 Email
 * @param {string} to - 收件人 Email
 * @param {string} subject - 主旨
 * @param {string} body - HTML 內容
 */
function sendEmailNotification(to, subject, body) {
  try {
    if (!to) return;
    GmailApp.sendEmail(to, subject, '', { htmlBody: body });
    logAction('SEND_EMAIL', 'N/A', `發送 Email 給 ${to}, 主旨: ${subject}`);
  } catch (error) {
    Logger.log(`發送 Email 失敗 (${to}): ${error.message}`);
  }
}

/**
 * 發送 WhatsApp 訊息 (Meta Cloud API)
 * @param {string} to - 收件人電話號碼 (需包含國碼，如 85298765432)
 * @param {string} message - 訊息內容
 */
function sendWhatsAppMessage(to, message) {
  try {
    if (!to) return;
    
    // 自動補齊香港國碼 (若只有 8 碼)
    let formatPhone = String(to).replace(/\D/g, '');
    if (formatPhone.length === 8) formatPhone = '852' + formatPhone;

    const configs = getSheetData(SHEET_NAMES.CONFIG);
    const tokenObj = configs.find(c => c['設定名稱'] === 'WHATSAPP_API_TOKEN' || c.key === 'WHATSAPP_API_TOKEN');
    const phoneIdObj = configs.find(c => c['設定名稱'] === 'WHATSAPP_PHONE_ID' || c.key === 'WHATSAPP_PHONE_ID');
    
    if (!tokenObj || !tokenObj.value || tokenObj.value === 'YOUR_TOKEN_HERE') {
      Logger.log('⚠️ WhatsApp 未發送：請先於 Config 設定 WHATSAPP_API_TOKEN');
      return;
    }

    const apiUrl = `https://graph.facebook.com/v17.0/${phoneIdObj.value}/messages`;
    
    const payload = {
      messaging_product: "whatsapp",
      recipient_type: "individual",
      to: formatPhone,
      type: "text",
      text: { body: message }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": `Bearer ${tokenObj.value}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    logAction('SEND_WHATSAPP', 'N/A', `發送 WhatsApp 給 ${formatPhone}, 狀態碼: ${response.getResponseCode()}`);
  } catch (error) {
    Logger.log(`發送 WhatsApp 失敗 (${to}): ${error.message}`);
  }
}

// ==========================================
// 3. 業務邏輯推播
// ==========================================

/**
 * 任務指派通知
 */
function notifyTaskAssigned(jobNumber, assignedToEmail, deadline) {
  const user = getUserByEmail(assignedToEmail);
  if (!user) return;

  // 1. 發送 Email
  const subject = `[新任務] 專案 ${jobNumber} 需要您的處理`;
  const body = `<h3>新任務指派</h3><p>您已被指派處理專案 <b>${jobNumber}</b>。</p><p>死線：${deadline}</p>`;
  sendEmailNotification(assignedToEmail, subject, body);

  // 2. 檢查是否為急件 (小於 3 個工作天)
  const today = new Date();
  const diffDays = getWorkingDaysDiff(today, deadline);
  
  if (diffDays < 3 && user.phoneNumber) {
    const waMsg = `🚨 [急件任務]\n專案 ${jobNumber} 已指派給您！\n距離死線 ${deadline} 剩餘不到 3 個工作天，請盡速處理。`;
    sendWhatsAppMessage(user.phoneNumber, waMsg);
  }
}

/**
 * 狀態變更通知
 */
function notifyStatusChange(jobNumber, oldStatus, newStatus) {
  const project = getProject(jobNumber);
  if (!project) return;

  const pmUser = getUserByName(project.pmName);
  const salesUser = getUserByName(project.salesPerson);
  const emailList = [];

  // 根據不同狀態決定通知對象 (簡化版：實際 Team Head 需依照部門邏輯擴充)
  if (newStatus === 'Waiting for PIC') {
    // 假設通知 PM 代表的 Team Head (可擴充)
    if (pmUser) emailList.push(pmUser.email);
  } else if (newStatus === 'Client Review' || newStatus === 'Completed') {
    if (pmUser) emailList.push(pmUser.email);
    if (salesUser) emailList.push(salesUser.email);
  } else if (newStatus === 'Blocked' || newStatus === 'Waiting Info') {
    if (pmUser) emailList.push(pmUser.email);
  }

  // 去除重複並發送
  const uniqueEmails = [...new Set(emailList)];
  uniqueEmails.forEach(email => {
    sendEmailNotification(email, `[狀態更新] 專案 ${jobNumber}`, `專案狀態已從 <b>${oldStatus}</b> 變更為 <b>${newStatus}</b>。`);
  });
}

/**
 * 每日任務推送 (Daily Digest) - 預計設定每日早上 9 點執行
 */
function sendDailyDigest() {
  const tasks = getSheetData(SHEET_NAMES.TASKS);
  // 過濾出尚未完成的任務
  const pendingTasks = tasks.filter(t => t.status !== 'Completed');
  
  // 依據 assignedTo 進行分組
  const tasksByUser = {};
  pendingTasks.forEach(t => {
    if (!tasksByUser[t.assignedTo]) tasksByUser[t.assignedTo] = [];
    tasksByUser[t.assignedTo].push(t);
  });

  for (const email in tasksByUser) {
    const user = getUserByEmail(email);
    const userTasks = tasksByUser[email];
    
    if (user) {
      let emailBody = `<h3>您的每日任務清單 (共 ${userTasks.length} 件)</h3><ul>`;
      let waMsg = `📋 [每日任務 Digest]\n早晨 ${user.name}！您今日有 ${userTasks.length} 件進行中任務：\n`;
      
      userTasks.forEach(t => {
        emailBody += `<li>專案: ${t.jobNumber} | 類型: ${t.taskType} | 狀態: ${t.status} | 死線: ${t.deadline}</li>`;
        waMsg += `- ${t.jobNumber} (${t.status})\n`;
      });
      emailBody += `</ul><p>請盡速跟進，謝謝！</p>`;
      
      sendEmailNotification(email, `[HK01 專案管理] 您的每日任務清單`, emailBody);
      if (user.phoneNumber) sendWhatsAppMessage(user.phoneNumber, waMsg);
    }
  }
  
  logAction('DAILY_DIGEST', 'ALL', `已發送每日推送給 ${Object.keys(tasksByUser).length} 位同事`);
}

// ==========================================
// 4. 測試程式碼
// ==========================================
function testNotification() {
  ensureWhatsAppConfig();
  
  const myEmail = Session.getActiveUser().getEmail();
  
  Logger.log('--- 1. 測試發送 Email ---');
  sendEmailNotification(myEmail, '測試 Email 功能', '<h1>這是一封測試信</h1><p>如果看到這封信，代表 Email 功能正常！</p>');
  Logger.log('Email 發送指令已執行，請檢查信箱。');
  
  Logger.log('--- 2. 測試狀態更新分發 ---');
  // 假定剛建立的 A26-1002 專案狀態變更
  notifyStatusChange('A26-1002', 'In Progress', 'Client Review');
  Logger.log('狀態更新分發已執行。');
  
  Logger.log('✅ NotificationLogic 測試完成！(WhatsApp 需填入真實 Token 才會生效)');
}