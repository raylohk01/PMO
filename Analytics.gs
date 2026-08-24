// ==========================================
// [Analytics.gs] 戰情室 API (新增 5 大部門「進行中 vs 即將進行」數據)
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
    const idxAudit = headers.findIndex(h => h === 'textjobtype' || h.includes('textjobtype') || h.includes('audit'));

    let kpi = { 
      overdue: 0, dueSoon: 0, onTrack: 0, unassigned: 0, 
      inClientReview: 0, next5DaysLaunch: 0, past5DaysLaunch: 0,
      completedCount: 0, onTimeCompletedCount: 0,
      totalClientDelayDays: 0, delayedClientReviewCount: 0,
      upcomingCount: 0, todayCreatedCount: 0, todayCompletedCount: 0  
    };

    // 💡 [新增] 5 大指定部門工作量計數器
    let deptWorkload = {
      Editorial: { active: 0, pipeline: 0 },
      Creative: { active: 0, pipeline: 0 },
      Design: { active: 0, pipeline: 0 },
      PM: { active: 0, pipeline: 0 },
      Video: { active: 0, pipeline: 0 }
    };

    let launchByDeptNext = { Editorial: 0, Creative: 0, Video: 0, Design: 0, Event: 0, Other: 0 };
    let launchByDeptPast = { Editorial: 0, Creative: 0, Video: 0, Design: 0, Event: 0, Other: 0 };
    let deptRevisionMap = {};
    let memberPerformance = {};
    let clientDelayList = [];

    let todayStart = new Date(now);
    todayStart.setHours(0,0,0,0);
    const next5DaysEnd = new Date(todayStart.getTime() + (5 * 24 * 60 * 60 * 1000));
    next5DaysEnd.setHours(23,59,59,999);
    const past5DaysStart = new Date(todayStart.getTime() - (5 * 24 * 60 * 60 * 1000));

    for (let i = 1; i < data.length; i++) {
      const jobNumber = idxJobNum >= 0 ? String(data[i][idxJobNum] || '').trim() : '';
      if (!jobNumber || jobNumber.toLowerCase() === 'jobnumber') continue;

      const pStatus = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim() : 'In Progress';
      if (pStatus === 'Recycle Bin' || pStatus === 'Cancelled') continue;

      const clientName = idxClient >= 0 ? String(data[i][idxClient] || '').trim() : '客戶';
      const pmName = idxPM >= 0 ? String(data[i][idxPM] || '').trim() : '';

      let isCreatedToday = false;
      if (idxAudit >= 0) {
        let auditStr = String(data[i][idxAudit] || '').trim();
        if (auditStr.includes('Create Project') && auditStr.includes(todayStr)) {
          isCreatedToday = true;
        }
      }

      let wfData = {};
      for (let c = 0; c < data[i].length; c++) {
        let cellStr = String(data[i][c] || '');
        if (cellStr.includes('deliverables')) { try { wfData = JSON.parse(cellStr); break; } catch(e){} }
      }

      if (wfData.deliverables) {
        const deliverableCount = wfData.deliverables.length; 

        wfData.deliverables.forEach((d, dIdx) => {
          if (d.status === 'Deleted' || d.status === 'Recycle Bin') return;
          
          const displayJobNum = (deliverableCount > 1 && jobNumber) ? `${jobNumber}-P${dIdx + 1}` : jobNumber;

          if (isCreatedToday) kpi.todayCreatedCount++;

          if (d.status === 'Completed' && d.completedAt && String(d.completedAt).includes(todayStr)) {
            kpi.todayCompletedCount++;
          }

          if (d.workflow) {
            let activeStep = d.workflow.find(s => s.status === 'In Progress');
            const allCompleted = d.status === 'Completed' || d.workflow.every(s => s.status === 'Completed');
            
            if (!activeStep && !allCompleted && d.workflow.length > 0) {
              activeStep = d.workflow.find(s => s.step === d.currentStep) || d.workflow[0];
            }

            if (!allCompleted && activeStep && (activeStep.status === 'Pending Start' || activeStep.status === 'Pending')) {
              kpi.upcomingCount++;
            }

            // 💡 統計 5 大部門「進行中」與「即將進行 (未來預估線)」
            if (!allCompleted) {
              d.workflow.forEach(s => {
                let sDeptKey = String(s.dept || '').trim();
                let matchedDept = Object.keys(deptWorkload).find(k => k.toLowerCase() === sDeptKey.toLowerCase());

                if (matchedDept) {
                  if (s.status === 'In Progress') {
                    deptWorkload[matchedDept].active++;
                  } else if (s.status === 'Pending' || s.status === 'Pending Start') {
                    if (!activeStep || s.step > activeStep.step || activeStep.status === 'Pending Start') {
                      deptWorkload[matchedDept].pipeline++;
                    }
                  }
                }
              });
            }

            let dLaunchDateStr = '';
            if (idxDeadline >= 0 && data[i][idxDeadline]) {
              let parsed = new Date(data[i][idxDeadline]);
              if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 2000) {
                dLaunchDateStr = Utilities.formatDate(parsed, "GMT+8", "yyyy-MM-dd");
              }
            }
            if (d.workflow.length > 0) {
              let lastStep = d.workflow[d.workflow.length - 1];
              if (lastStep.keyDate || lastStep.deadline) {
                dLaunchDateStr = lastStep.keyDate || lastStep.deadline;
              }
            }

            if (dLaunchDateStr) {
              let lDate = new Date(dLaunchDateStr);
              lDate.setHours(0,0,0,0);
              let mainDept = d.workflow[0] ? d.workflow[0].dept || 'Other' : 'Other';

              if (lDate >= todayStart && lDate <= next5DaysEnd) {
                kpi.next5DaysLaunch++;
                launchByDeptNext[mainDept] = (launchByDeptNext[mainDept] || 0) + 1;
              }
              if (lDate >= past5DaysStart && lDate <= todayStart) {
                kpi.past5DaysLaunch++;
                launchByDeptPast[mainDept] = (launchByDeptPast[mainDept] || 0) + 1;
              }
            }

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

              if (activeStep && s.step === activeStep.step && !allCompleted) {
                const effDate = s.keyDate || todayStr;
                const taskDate = new Date(effDate);

                if (s.status === 'In Progress') {
                  if (effDate < todayStr) kpi.overdue++;
                  else if (!s.assignee || s.assignee === '未指派') kpi.unassigned++;
                  else if (taskDate <= new Date(now.getTime() + (48 * 60 * 60 * 1000))) kpi.dueSoon++;
                  else kpi.onTrack++;
                }

                if (dept.toLowerCase().includes('client') || s.name.toLowerCase().includes('client review') || s.name.includes('客戶審批')) {
                  kpi.inClientReview++;
                  if (effDate < todayStr && s.status === 'In Progress') {
                    let delayMs = now.getTime() - taskDate.getTime();
                    let delayDays = Math.ceil(delayMs / (1000 * 60 * 60 * 24));
                    kpi.totalClientDelayDays += delayDays;
                    kpi.delayedClientReviewCount++;

                    clientDelayList.push({
                      jobNumber: displayJobNum,
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
        deptWorkload: deptWorkload, // 💡 回傳部門工作量數據
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
      delayedClientReviewCount: 0,
      upcomingCount: 0,       // 💡
      todayCreatedCount: 0,   // 💡
      todayCompletedCount: 0  // 💡
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