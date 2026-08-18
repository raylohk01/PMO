/**
 * 📄 ReportLogic.gs
 * 績效報表與數據分析 API (供前端 Chart.js 畫圖使用)
 */

// ==========================================
// 1. 計算逾期率 (Late Rate)
// ==========================================
/**
 * 計算專案逾期率
 * @param {string} startDate - YYYY-MM-DD
 * @param {string} endDate - YYYY-MM-DD
 * @returns {Object} JSON 報表數據
 */
function calculateLateRate(startDate, endDate) {
  try {
    const projects = getSheetData(SHEET_NAMES.PROJECTS);
    
    // 篩選出已完成的專案，並過濾日期範圍 (如果有傳入)
    const completedProjects = projects.filter(p => {
      if (p.status !== 'Completed') return false;
      if (!p.completedDate) return false;
      
      const compDate = new Date(p.completedDate);
      if (startDate && compDate < new Date(startDate)) return false;
      if (endDate && compDate > new Date(endDate)) return false;
      return true;
    });

    const pmStats = {};
    let totalCompleted = 0;
    let totalLate = 0;

    completedProjects.forEach(p => {
      const pm = p.pmName || 'Unassigned';
      if (!pmStats[pm]) pmStats[pm] = { completed: 0, late: 0 };
      
      pmStats[pm].completed++;
      totalCompleted++;
      
      // 逾期判斷：completedDate > firstDraftDate
      if (p.firstDraftDate) {
        const compDate = new Date(p.completedDate).setHours(0,0,0,0);
        const draftDate = new Date(p.firstDraftDate).setHours(0,0,0,0);
        if (compDate > draftDate) {
          pmStats[pm].late++;
          totalLate++;
        }
      }
    });

    // 格式化為前端圖表易用的陣列
    const pmData = Object.keys(pmStats).map(pm => {
      const rate = pmStats[pm].completed === 0 ? 0 : (pmStats[pm].late / pmStats[pm].completed) * 100;
      return { pmName: pm, lateRate: parseFloat(rate.toFixed(1)), lateCount: pmStats[pm].late, totalCount: pmStats[pm].completed };
    });

    const overallRate = totalCompleted === 0 ? 0 : (totalLate / totalCompleted) * 100;

    return { 
      success: true, 
      data: { 
        overallLateRate: parseFloat(overallRate.toFixed(1)), 
        pmBreakdown: pmData 
      } 
    };
  } catch (error) {
    return { success: false, message: '計算逾期率失敗: ' + error.message };
  }
}

// ==========================================
// 2. 高耗損客戶分析 (High Churn Clients)
// ==========================================
/**
 * 找出退件次數 (versionNo) >= 5 的專案與客戶
 * @returns {Object} JSON 報表數據
 */
function getHighChurnClients() {
  try {
    const tasks = getSheetData(SHEET_NAMES.TASKS);
    const projects = getSheetData(SHEET_NAMES.PROJECTS);
    
    // 找出所有退件次數過高的任務 (這在 Task 表中記錄)
    const highRevisionTasks = tasks.filter(t => (parseInt(t.versionNo) || 0) >= 5);
    
    const clientStats = {};
    
    highRevisionTasks.forEach(t => {
      const proj = projects.find(p => p.jobNumber === t.jobNumber) || {};
      const client = proj.clientName || '未知客戶';
      const vNo = parseInt(t.versionNo) || 0;
      
      if (!clientStats[client]) {
        clientStats[client] = { totalRevisions: 0, projects: new Set() };
      }
      
      clientStats[client].totalRevisions += vNo;
      clientStats[client].projects.add(t.jobNumber);
    });
    
    // 轉換為陣列並排序
    const reportData = Object.keys(clientStats).map(client => {
      return {
        clientName: client,
        totalRevisions: clientStats[client].totalRevisions,
        involvedProjectsCount: clientStats[client].projects.size,
        involvedProjects: Array.from(clientStats[client].projects)
      };
    });
    
    // 按退件總數由高至低排序
    reportData.sort((a, b) => b.totalRevisions - a.totalRevisions);
    
    return { success: true, data: reportData };
  } catch (error) {
    return { success: false, message: '獲取高耗損客戶分析失敗: ' + error.message };
  }
}

// ==========================================
// 3. 各部門效率 (Department Performance)
// ==========================================
/**
 * 統計部門總任務數、逾期率、平均退件數
 * @returns {Object} JSON 報表數據
 */
function getDepartmentPerformance() {
  try {
    const tasks = getSheetData(SHEET_NAMES.TASKS);
    const users = getSheetData(SHEET_NAMES.USERS);
    
    const deptStats = {};
    
    tasks.forEach(t => {
      const user = users.find(u => u.email === t.assignedTo);
      const dept = user ? user.department : 'Unknown';
      
      if (!deptStats[dept]) {
        deptStats[dept] = { totalTasks: 0, lateTasks: 0, totalVersions: 0, completedTasks: 0, totalDaysSpent: 0 };
      }
      
      deptStats[dept].totalTasks++;
      deptStats[dept].totalVersions += (parseInt(t.versionNo) || 0);
      
      if (t.status === 'Completed') {
        deptStats[dept].completedTasks++;
        // 計算任務逾期
        if (t.deadline) {
           const compDate = new Date().setHours(0,0,0,0); // 簡化：假設今日完成，實際應查 Log，這裡用 deadline vs today 模擬
           const dlDate = new Date(t.deadline).setHours(0,0,0,0);
           if (compDate > dlDate) deptStats[dept].lateTasks++;
        }
      }
    });
    
    const reportData = Object.keys(deptStats).map(dept => {
      const stats = deptStats[dept];
      const avgVersions = stats.totalTasks === 0 ? 0 : stats.totalVersions / stats.totalTasks;
      const lateRate = stats.completedTasks === 0 ? 0 : (stats.lateTasks / stats.completedTasks) * 100;
      
      return {
        department: dept,
        totalTasks: stats.totalTasks,
        avgRevisionCount: parseFloat(avgVersions.toFixed(1)),
        lateRate: parseFloat(lateRate.toFixed(1))
      };
    });
    
    return { success: true, data: reportData };
  } catch (error) {
    return { success: false, message: '獲取部門效率失敗: ' + error.message };
  }
}

// ==========================================
// 4. PM 工作量 (PM Workload)
// ==========================================
/**
 * 統計每個 PM 的 Active 專案數量與任務密度
 * @returns {Object} JSON 報表數據
 */
function getPMWorkload() {
  try {
    const projects = getSheetData(SHEET_NAMES.PROJECTS);
    const tasks = getSheetData(SHEET_NAMES.TASKS);
    
    const pmStats = {};
    
    projects.forEach(p => {
      const pm = p.pmName || 'Unassigned';
      if (!pmStats[pm]) {
        pmStats[pm] = { activeProjects: 0, totalTasksInActive: 0 };
      }
      
      if (p.status !== 'Completed') {
        pmStats[pm].activeProjects++;
        
        // 計算這個活躍專案底下有幾個任務
        const relatedTasks = tasks.filter(t => t.jobNumber === p.jobNumber).length;
        pmStats[pm].totalTasksInActive += relatedTasks;
      }
    });
    
    const reportData = Object.keys(pmStats).map(pm => {
      const stats = pmStats[pm];
      const avgTasks = stats.activeProjects === 0 ? 0 : stats.totalTasksInActive / stats.activeProjects;
      
      return {
        pmName: pm,
        activeProjects: stats.activeProjects,
        avgTasksPerProject: parseFloat(avgTasks.toFixed(1))
      };
    });
    
    // 按活躍專案數排序 (最忙的排前面)
    reportData.sort((a, b) => b.activeProjects - a.activeProjects);
    
    return { success: true, data: reportData };
  } catch (error) {
    return { success: false, message: '獲取 PM 工作量失敗: ' + error.message };
  }
}

// ==========================================
// 5. 測試程式碼
// ==========================================
function testReportLogic() {
  Logger.log('--- 1. 計算逾期率 (Late Rate) ---');
  const lateRate = calculateLateRate();
  Logger.log(JSON.stringify(lateRate, null, 2));
  
  Logger.log('--- 2. 高耗損客戶分析 (High Churn) ---');
  const churn = getHighChurnClients();
  Logger.log(JSON.stringify(churn, null, 2));
  
  Logger.log('--- 3. 測試部門效率 ---');
  const dept = getDepartmentPerformance();
  Logger.log(JSON.stringify(dept, null, 2));
  
  Logger.log('--- 4. 測試 PM 工作量 ---');
  const pmWorkload = getPMWorkload();
  Logger.log(JSON.stringify(pmWorkload, null, 2));
  
  Logger.log('✅ 績效報表測試完成！');
}