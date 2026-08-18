// ==========================================
// [Analytics.gs] 投屏戰情室全域數據分析 API (含前後 5 天上線統計)
// ==========================================

function api_getAnalyticsData(timeRange, customStart, customEnd) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const currentUser = getUserByEmail(userEmail);

    if (!currentUser) throw new Error('無法驗證使用者身份');

    const userRole = currentUser.role || 'Member';
    const userDept = currentUser.department || 'General';
    const userName = currentUser.name || userEmail.split('@')[0];

    const isSuperManager = ['Admin', 'Management'].includes(userRole) || userRole === 'Head of PM';
    const isTeamHead = userRole === 'Team Head';

    const now = new Date();
    const todayStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

    let startDate = new Date();
    let endDate = new Date();

    if (timeRange === 'TODAY' || !timeRange) {
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(23, 59, 59, 999);
    } else if (timeRange === 'LAST_5_DAYS') {
      startDate.setDate(now.getDate() - 5);
      startDate.setHours(0, 0, 0, 0);
    } else if (timeRange === 'NEXT_14_DAYS') {
      endDate.setDate(now.getDate() + 14);
      endDate.setHours(23, 59, 59, 999);
    } else if (timeRange === 'LAST_30_DAYS') {
      startDate.setDate(now.getDate() - 30);
      startDate.setHours(0, 0, 0, 0);
    } else if (timeRange === 'CUSTOM' && customStart && customEnd) {
      startDate = new Date(customStart);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(customEnd);
      endDate.setHours(23, 59, 59, 999);
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    if (!sheet) return { success: true, data: getEmptyAnalyticsResult() };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: getEmptyAnalyticsResult() };

    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const idxJobNum = headers.findIndex(h => h.includes('jobnumber') || h === 'jobno');
    const idxClient = headers.findIndex(h => h.includes('client'));
    const idxPM = headers.findIndex(h => h.includes('pmname') || h === 'pm');
    const idxStatus = headers.findIndex(h => h === 'status' || h === 'project_status');
    const idxDeadline = headers.findIndex(h => h.includes('launch') || h.includes('deadline'));

    let kpi = { 
      overdue: 0, dueSoon: 0, onTrack: 0, unassigned: 0, 
      inClientReview: 0, next5DaysLaunch: 0, past5DaysLaunch: 0,
      completedCount: 0, onTimeCompletedCount: 0,
      totalClientDelayDays: 0, delayedClientReviewCount: 0
    };

    let launchByDeptNext = { Editorial: 0, Creative: 0, Video: 0, Design: 0, Event: 0, Other: 0 };
    let launchByDeptPast = { Editorial: 0, Creative: 0, Video: 0, Design: 0, Event: 0, Other: 0 };
    let deptRevisionMap = {};
    let memberPerformance = {};
    let clientDelayList = [];

    const fortyEightHoursLater = new Date(now.getTime() + (48 * 60 * 60 * 1000));
    const next5DaysEnd = new Date(now.getTime() + (5 * 24 * 60 * 60 * 1000));
    const past5DaysStart = new Date(now.getTime() - (5 * 24 * 60 * 60 * 1000));

    for (let i = 1; i < data.length; i++) {
      const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (!jobNumber || jobNumber.toLowerCase() === 'jobnumber') continue;

      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : 'In Progress';
      if (pStatus === 'Recycle Bin' || pStatus === 'Cancelled') continue;

      const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '客戶';
      const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';

      let wfData = {};
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables')) { try { wfData = JSON.parse(cellStr); break; } catch(e){} }
      }

      // 💡 掃描專案上線日期（過去 5 天 vs 未來 5 天）
      if (idxDeadline >= 0 && data[i][idxDeadline]) {
        let launchDate = new Date(data[i][idxDeadline]);
        if (!isNaN(launchDate.getTime())) {
          let mainDept = 'Other';
          if (wfData.deliverables && wfData.deliverables.length > 0) {
            let firstWorkflow = wfData.deliverables[0].workflow;
            if (firstWorkflow && firstWorkflow.length > 0) mainDept = firstWorkflow[0].dept || 'Other';
          }

          // 未來 5 天預計上線
          if (launchDate >= now && launchDate <= next5DaysEnd) {
            kpi.next5DaysLaunch++;
            if (launchByDeptNext[mainDept] !== undefined) launchByDeptNext[mainDept]++;
            else launchByDeptNext['Other']++;
          }

          // 過去 5 天已上線
          if (launchDate >= past5DaysStart && launchDate <= now) {
            kpi.past5DaysLaunch++;
            if (launchByDeptPast[mainDept] !== undefined) launchByDeptPast[mainDept]++;
            else launchByDeptPast['Other']++;
          }
        }
      }

      // 關卡與客戶審批掃描
      if (wfData.deliverables) {
        wfData.deliverables.forEach(d => {
          if (d.workflow) {
            d.workflow.forEach(s => {
              const dept = s.dept || 'Other';
              const assignee = s.assignee || '未指派';

              let allowData = false;
              if (isSuperManager) allowData = true;
              else if (isTeamHead && dept.toLowerCase() === userDept.toLowerCase()) allowData = true;
              else if (assignee.toLowerCase() === userName.toLowerCase() || pmName.toLowerCase() === userName.toLowerCase()) allowData = true;

              if (!allowData) return;

              if (s.revisionCount && s.revisionCount > 0) {
                deptRevisionMap[dept] = (deptRevisionMap[dept] || 0) + s.revisionCount;
              }

              if (s.status === 'Completed' && s.completedAt) {
                let compDate = new Date(s.completedAt);
                if (compDate >= startDate && compDate <= endDate) {
                  kpi.completedCount++;
                  let effDeadline = s.keyDate ? new Date(s.keyDate) : null;
                  let isOnTime = !effDeadline || compDate <= effDeadline;
                  if (isOnTime) kpi.onTimeCompletedCount++;

                  if (!memberPerformance[assignee]) memberPerformance[assignee] = { name: assignee, dept: dept, totalTasks: 0, completedOnTime: 0, revisions: 0 };
                  memberPerformance[assignee].totalTasks++;
                  if (isOnTime) memberPerformance[assignee].completedOnTime++;
                }
              }

              if (s.status === 'In Progress') {
                const effDate = s.keyDate || todayStr;
                const taskDate = new Date(effDate);

                if (effDate < todayStr) kpi.overdue++;
                else if (!s.assignee) kpi.unassigned++;
                else if (taskDate <= fortyEightHoursLater) kpi.dueSoon++;
                else kpi.onTrack++;

                if (dept.toLowerCase().includes('client') || s.name.toLowerCase().includes('client review') || s.name.includes('客戶審批')) {
                  kpi.inClientReview++;
                  if (effDate < todayStr) {
                    let delayMs = now.getTime() - taskDate.getTime();
                    let delayDays = Math.ceil(delayMs / (1000 * 60 * 60 * 24));
                    kpi.totalClientDelayDays += delayDays;
                    kpi.delayedClientReviewCount++;

                    clientDelayList.push({
                      jobNumber: jobNumber,
                      client: clientName,
                      pmName: pmName,
                      taskName: d.name || '項目',
                      delayDays: delayDays,
                      deadline: effDate
                    });
                  }
                }

                if (assignee !== '未指派') {
                  if (!memberPerformance[assignee]) memberPerformance[assignee] = { name: assignee, dept: dept, totalTasks: 0, completedOnTime: 0, revisions: 0 };
                  if (s.revisionCount) memberPerformance[assignee].revisions += s.revisionCount;
                }
              }
            });
          }
        });
      }
    }

    const onTimeRate = kpi.completedCount > 0 ? Math.round((kpi.onTimeCompletedCount / kpi.completedCount) * 100) : 100;
    const avgClientDelayDays = kpi.delayedClientReviewCount > 0 ? (kpi.totalClientDelayDays / kpi.delayedClientReviewCount).toFixed(1) : 0;

    return {
      success: true,
      data: {
        timeRange: timeRange,
        lastRefreshed: Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss"),
        kpi: kpi,
        onTimeRate: onTimeRate,
        avgClientDelayDays: avgClientDelayDays,
        launchByDeptNext: launchByDeptNext,
        launchByDeptPast: launchByDeptPast,
        deptRevisionMap: deptRevisionMap,
        memberPerformance: Object.values(memberPerformance),
        clientDelayList: clientDelayList
      }
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

// 🛡️ 靜態防護：當資料庫無資料時回傳乾淨且結構完整的空物件
function getEmptyAnalyticsResult() {
  return {
    kpi: { 
      overdue: 0, 
      dueSoon: 0, 
      onTrack: 0, 
      unassigned: 0, 
      inClientReview: 0, 
      next5DaysLaunch: 0, 
      past5DaysLaunch: 0,
      completedCount: 0,
      onTimeCompletedCount: 0,
      totalClientDelayDays: 0,
      delayedClientReviewCount: 0
    },
    onTimeRate: 100,
    avgClientDelayDays: 0,
    launchByDeptNext: { Editorial: 0, Creative: 0, Video: 0, Design: 0, Event: 0, Other: 0 },
    launchByDeptPast: { Editorial: 0, Creative: 0, Video: 0, Design: 0, Event: 0, Other: 0 },
    deptRevisionMap: {},
    memberPerformance: [],
    clientDelayList: []
  };
}

function test_Defensive_ManageProject() {
  // 故意傳入一個絕對不存在的 ID
  let result = api_manageProjectStatus('FAKE_JOB_9999', 'PAUSE');
  Logger.log(result.message);
}