/**
 * 📄 ViewLogic.gs
 * Project Overview 三層視圖資料 API
 * 供前端 Web App (HTML Service) 透過 google.script.run 呼叫
 */

// ==========================================
// 1. 執行同事視圖 (Individual View)
// ==========================================
/**
 * 獲取個人任務清單，按緊急程度排序
 * @param {string} userEmail 
 * @returns {Object} JSON 格式的分組任務資料
 */
function getIndividualView(userEmail) {
  try {
    const allTasks = getSheetData(SHEET_NAMES.TASKS);
    const allProjects = getSheetData(SHEET_NAMES.PROJECTS);

    // 過濾出分配給自己且未完成的任務
    let myTasks = allTasks.filter(t => t.assignedTo === userEmail && t.status !== 'Completed');

    // 豐富資料：加入專案資訊與 deadline 狀態
    const enrichedTasks = myTasks.map(t => {
       const proj = allProjects.find(p => p.jobNumber === t.jobNumber) || {};
       
       const today = new Date();
       today.setHours(0, 0, 0, 0);
       const deadline = new Date(t.deadline);
       deadline.setHours(0, 0, 0, 0);

       let deadlineStatus = 'onTrack';
       if (today > deadline) {
         deadlineStatus = 'overdue';
       } else {
         const diff = getWorkingDaysDiff(today, deadline);
         if (diff < 3) deadlineStatus = 'urgent';
       }

       return {
         taskId: t.taskId,
         jobNumber: t.jobNumber,
         client: proj.clientName || '未知客戶',
         taskType: t.taskType || '未分類',
         status: t.status,
         deadline: typeof t.deadline === 'string' ? t.deadline : formatDate(t.deadline),
         deadlineStatus: deadlineStatus,
         dateObj: deadline.getTime() // 用於排序
       };
    });

    // 按照死線由近到遠排序
    enrichedTasks.sort((a, b) => a.dateObj - b.dateObj);

    // 分組
    const pending = enrichedTasks.filter(t => t.status === 'Waiting for PIC');
    const inProgress = enrichedTasks.filter(t => ['In Progress', 'Waiting for First Draft', 'Client Review', 'Internal Review'].includes(t.status));
    const upcoming = enrichedTasks.filter(t => !['Waiting for PIC', 'In Progress', 'Waiting for First Draft', 'Client Review', 'Internal Review'].includes(t.status));

    return { 
      success: true, 
      data: { pending, inProgress, upcoming } 
    };
  } catch (error) {
    return { success: false, message: '讀取個人視圖失敗: ' + error.message };
  }
}

// ==========================================
// 2. 部門主管視圖 (Department Head View)
// ==========================================
/**
 * 獲取部門整體工作量與進度
 * @param {string} department - 如 'Editorial', 'Creative'
 * @returns {Object} 
 */
function getDepartmentHeadView(department) {
  try {
    const allTasks = getSheetData(SHEET_NAMES.TASKS);
    const allProjects = getSheetData(SHEET_NAMES.PROJECTS);
    const allUsers = getSheetData(SHEET_NAMES.USERS);

    // 找出該部門的所有同事
    const deptMembers = allUsers.filter(u => String(u.department).toUpperCase() === String(department).toUpperCase());
    const deptEmails = deptMembers.map(u => u.email);

    const deptTasks = allTasks.filter(t => deptEmails.includes(t.assignedTo) && t.status !== 'Completed');

    // 初始化同事工作量矩陣
    const workload = {};
    deptMembers.forEach(user => {
      workload[user.name] = { count: 0, overdue: 0, urgent: 0 };
    });

    const inProgress = [];
    const waitingForPIC = [];
    const outstanding = [];

    deptTasks.forEach(t => {
      const proj = allProjects.find(p => p.jobNumber === t.jobNumber) || {};
      const user = allUsers.find(u => u.email === t.assignedTo);
      const userName = user ? user.name : t.assignedTo;

      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const deadline = new Date(t.deadline);
      deadline.setHours(0, 0, 0, 0);

      const isOverdue = today > deadline;
      const isUrgent = !isOverdue && getWorkingDaysDiff(today, deadline) < 3;

      const taskData = {
         taskId: t.taskId,
         jobNumber: t.jobNumber,
         client: proj.clientName || '未知',
         PIC: userName,
         status: t.status,
         deadline: typeof t.deadline === 'string' ? t.deadline : formatDate(t.deadline)
      };

      // 裝填到不同視圖桶子
      if (isOverdue) {
        outstanding.push(taskData);
      } else if (t.status === 'Waiting for PIC') {
        waitingForPIC.push(taskData);
      } else {
        inProgress.push(taskData);
      }

      // 更新負載統計
      if (workload[userName]) {
        workload[userName].count++;
        if (isOverdue) workload[userName].overdue++;
        if (isUrgent) workload[userName].urgent++;
      }
    });

    return { 
      success: true, 
      data: { workload, waitingForPIC, inProgress, outstanding } 
    };
  } catch (error) {
     return { success: false, message: '讀取部門視圖失敗: ' + error.message };
  }
}

// ==========================================
// 3. 全局管理層視圖 (Management View)
// ==========================================
/**
 * 獲取全公司概況與高風險專案
 * @returns {Object}
 */
function getManagementView() {
  try {
    const allTasks = getSheetData(SHEET_NAMES.TASKS);
    const allProjects = getSheetData(SHEET_NAMES.PROJECTS);
    const allUsers = getSheetData(SHEET_NAMES.USERS);

    const deptStats = {};
    const departments = ['PM', 'Editorial', 'Creative', 'Art', 'Event', 'Sales'];
    departments.forEach(d => {
       deptStats[d] = { totalActive: 0, blocked: 0, overdue: 0 };
    });

    const highRiskProjects = [];
    const seenHighRisk = new Set();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    allTasks.forEach(t => {
      if (t.status === 'Completed') return;

      const user = allUsers.find(u => u.email === t.assignedTo);
      const dept = user ? user.department : 'Unknown';
      
      const deadline = new Date(t.deadline);
      deadline.setHours(0, 0, 0, 0);

      let overdueDays = 0;
      if (today > deadline) {
         overdueDays = getWorkingDaysDiff(deadline, today); // 計算過期幾日
      }

      // 部門統計累加
      if (deptStats[dept]) {
         deptStats[dept].totalActive++;
         if (t.status === 'Blocked') deptStats[dept].blocked++;
         if (overdueDays > 0) deptStats[dept].overdue++;
      }

      // 篩選高風險專案 (Blocked 或逾期超過 5 個工作天)
      if (t.status === 'Blocked' || overdueDays >= 5) {
         if (!seenHighRisk.has(t.jobNumber)) {
            seenHighRisk.add(t.jobNumber);
            const proj = allProjects.find(p => p.jobNumber === t.jobNumber) || {};
            highRiskProjects.push({
               jobNumber: t.jobNumber,
               client: proj.clientName || '未知',
               reason: t.status === 'Blocked' ? 'Blocked (卡關中)' : `嚴重逾期 (${overdueDays} 個工作日)`,
               PIC: user ? user.name : t.assignedTo
            });
         }
      }
    });

    return { 
      success: true, 
      data: { deptStats, highRiskProjects } 
    };
  } catch (error) {
     return { success: false, message: '讀取全局視圖失敗: ' + error.message };
  }
}

// ==========================================
// 4. 測試 API 輸出格式
// ==========================================
function testViewsAPI() {
  const myEmail = Session.getActiveUser().getEmail();
  
  Logger.log('--- 1. 測試個人視圖 (Individual) ---');
  const individual = getIndividualView(myEmail); // 或替換成您指派過任務的 Email
  Logger.log(JSON.stringify(individual, null, 2));
  
  Logger.log('--- 2. 測試部門視圖 (Department Head) ---');
  // 假設剛才您測試的指派給了 ming@hk01.com，他是 Editorial 部門
  const dept = getDepartmentHeadView('Editorial');
  Logger.log(JSON.stringify(dept, null, 2));
  
  Logger.log('--- 3. 測試全局視圖 (Management) ---');
  const mgmt = getManagementView();
  Logger.log(JSON.stringify(mgmt, null, 2));
  
  Logger.log('✅ View API 測試完成！請檢視輸出的 JSON 結構。');
}