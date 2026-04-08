// ============================================================
// Global Analytics Utilities
// ============================================================
function isRecurringRow_(row) {
  var tags = String(row[22] || '');
  if (!tags) return false;
  return tags.split(',').map(function (t) { return String(t).trim(); }).indexOf('__recurring') !== -1;
}

// ------------------------------------------------------------------

// ------------------------------------------------------------------
// CACHE KEY HELPER
// ------------------------------------------------------------------
function dashCacheKey_(filters) {
  var f = filters || {};
  var version = PropertiesService.getScriptProperties().getProperty('DASH_CACHE_VERSION') || '1';
  return 'dash_v' + version + '_' + (f.project || '') + '_' + (f.team || '') + '_' + (f.member || '') + '_' + (f.type || '') + '_' + (f.from || '') + '_' + (f.to || '');
}

// ------------------------------------------------------------------
// SINGLE-PASS TASK AGGREGATOR
// Returns raw buckets used by all sub-functions. Called once per
// getAnalyticsDashboard request — all metric functions receive the
// result rather than re-reading the sheet.
// ------------------------------------------------------------------
function aggregateTasks_(filters) {
  filters = filters || {};
  var taskSheet = getSheet(SHEETS.TASKS);
  var allRows = taskSheet.getDataRange().getValues();
  var memberMap = getMemberMap_();
  var now = new Date();
  var todayStr = now.toISOString().substring(0, 10);

  var projectNameMap = {};
  try {
    if (typeof getProjectsPayload_ === 'function') {
      var actor = requireRole_(['Owner', 'Manager']);
      var projectRows = getProjectsPayload_(actor.email, actor.role) || [];
      projectRows.forEach(function (p) {
        if (p && p.projectId) projectNameMap[p.projectId] = p.name || p.projectId;
      });
    }
  } catch (e) { }

  var fromDate = filters.from ? new Date(filters.from) : getDateDaysAgo_(30);
  var toDate = filters.to ? new Date(filters.to) : now;
  toDate.setHours(23, 59, 59, 999);

  var weekStart = new Date(now);
  weekStart.setDate(weekStart.getDate() - ((weekStart.getDay() + 6) % 7));
  weekStart.setHours(0, 0, 0, 0);

  var summary = { tasksCreatedToday: 0, tasksCompletedToday: 0, tasksCompletedThisWeek: 0, overdueTasks: 0, activeTasks: 0, totalInWindow: 0, bottleneckTeam: 'None' };
  var globalSummary = { tasksCreatedToday: 0, tasksCompletedToday: 0, tasksCompletedThisWeek: 0, overdueTasks: 0, activeTasks: 0, totalInWindow: 0 };
  var globalByTeam = {};

  var totalCompletionHours = 0, completedWithTAT = 0;
  var globalTotalCompletionHours = 0, globalCompletedWithTAT = 0;

  var byUser = {};
  var byTeam = {};
  var byStatus = {};
  var byType = {};
  var byPriority = {};
  var byProject = {};
  var trendMap = {};
  var detailedTasks = [];
  var allTaskMeta = {};
  var statusAging = {}; // Bottleneck Analysis
  var teamSlaRisk = {}; // Risk Detection
  var slaConfig = getSLAConfig_();
  var teamCapacity = {};

  // Pre-calculate team capacity from active member count
  Object.values(memberMap).forEach(function (m) {
    if (m.active && m.team) teamCapacity[m.team] = (teamCapacity[m.team] || 0) + 1;
  });

  function ub(email, name, team) {
    var k = String(email || '').toLowerCase();
    if (!byUser[k]) byUser[k] = { name: name || k, email: k, team: team || '', active: 0, completed: 0, overdue: 0, breaches: 0, totalTAT: 0, typeStats: {} };
    return byUser[k];
  }
  function tb(team) {
    var key = team || 'Unknown';
    if (!byTeam[key]) byTeam[key] = { team: key, active: 0, completed: 0, overdue: 0, breaches: 0, totalTAT: 0, tatCount: 0, typeStats: {} };
    return byTeam[key];
  }
  function pb(projectId, projectName, team) {
    var key = projectId || 'NO_PROJECT';
    if (!byProject[key]) {
      byProject[key] = { projectId: key, projectName: projectName || 'No Project', team: team || 'Mixed', total: 0, active: 0, completed: 0, overdue: 0, breaches: 0 };
    }
    return byProject[key];
  }
  function bumpTrend(dateStr, field, val) {
    if (!trendMap[dateStr]) trendMap[dateStr] = { created: 0, completed: 0, tatHoursTotal: 0, tatCount: 0 };
    if (field === 'tat') {
      trendMap[dateStr].tatHoursTotal += val;
      trendMap[dateStr].tatCount++;
    } else {
      trendMap[dateStr][field]++;
    }
  }

  for (var i = 1; i < allRows.length; i++) {
    var row = allRows[i];
    if (!row[0]) continue;
    if (isRecurringRow_(row)) continue;

    var ownerEmail = String(row[COL.OWNER_EMAIL] || '').toLowerCase();
    var ownerName = row[COL.OWNER_NAME] || ownerEmail;
    var team = row[COL.CURRENT_TEAM] || 'Unknown';
    var status = row[COL.STATUS] || 'To Do';
    var createdAt = row[COL.CREATED_AT] ? new Date(row[COL.CREATED_AT]) : null;
    var deadline = row[COL.DEADLINE] ? new Date(row[COL.DEADLINE]) : null;
    var completedAt = row[COL.COMPLETED_AT] ? new Date(row[COL.COMPLETED_AT]) : null;
    var totalHrs = Number(row[COL.TOTAL_HOURS]) || 0;
    var breached = row[COL.SLA_BREACHED] === true;
    var projectId = row[COL.PROJECT_ID] || '';
    var projectName = projectId ? (projectNameMap[projectId] || projectId) : 'No Project';
    var taskId = row[COL.TASK_ID] || '';

    if (taskId) {
      allTaskMeta[taskId] = {
        taskId: taskId,
        taskName: row[COL.TASK_NAME] || taskId,
        projectId: projectId,
        projectName: projectName,
        ownerEmail: ownerEmail,
        ownerName: ownerName,
        team: team,
        createdAt: createdAt ? createdAt.toISOString() : '',
        completedAt: completedAt ? completedAt.toISOString() : ''
      };
    }

    // GLOBAL-ONLY Stats (Ignore Filters)
    var isDoneG = status === 'Done' || status === 'Completed' || status === 'Archived';
    var isOverdueG = !isDoneG && deadline && deadline < now;
    var isActiveG = !isDoneG && status !== 'On Hold';
    var inWindowG = false;
    if (completedAt) inWindowG = completedAt >= fromDate && completedAt <= toDate;
    else if (createdAt) inWindowG = createdAt >= fromDate && createdAt <= toDate;

    if (inWindowG) {
      globalSummary.totalInWindow++;
      if (createdAt && createdAt.toISOString().substring(0, 10) == todayStr) globalSummary.tasksCreatedToday++;
      if (completedAt && completedAt.toISOString().substring(0, 10) == todayStr) globalSummary.tasksCompletedToday++;
      if (completedAt && completedAt >= weekStart) globalSummary.tasksCompletedThisWeek++;
      if (isOverdueG) globalSummary.overdueTasks++;
      if (isActiveG) globalSummary.activeTasks++;

      if (!globalByTeam[team]) globalByTeam[team] = { overdueTasks: 0, activeTasks: 0, totalTAT: 0, count: 0, completedCount: 0 };
      if (isOverdueG) globalByTeam[team].overdueTasks++;
      if (isActiveG) globalByTeam[team].activeTasks++;
      if (isDoneG && totalHrs > 0) {
        globalByTeam[team].totalTAT += totalHrs;
        globalByTeam[team].count++;
        globalTotalCompletionHours += totalHrs;
        globalCompletedWithTAT++;
      }
    }

    // FILTER LOGIC
    if (filters.project && projectId !== filters.project) continue;
    if (filters.team && team !== filters.team) continue;
    if (filters.member && ownerEmail !== String(filters.member || '').toLowerCase()) continue;
    if (filters.type && row[COL.TASK_TYPE] !== filters.type) continue;

    var inWindow = false;
    if (completedAt) inWindow = completedAt >= fromDate && completedAt <= toDate;
    else if (createdAt) inWindow = createdAt >= fromDate && createdAt <= toDate;
    if (!inWindow) continue;

    summary.totalInWindow++;

    var isDone = status === 'Done' || status === 'Completed' || status === 'Archived';
    var isArchived = status === 'Archived';
    var isActive = !isDone && status !== 'On Hold';
    var isOverdue = !isDone && deadline && deadline < now;

    if (createdAt && createdAt.toISOString().substring(0, 10) == todayStr) summary.tasksCreatedToday++;
    if (completedAt && completedAt.toISOString().substring(0, 10) == todayStr) summary.tasksCompletedToday++;

    if (!isDone && deadline && deadline < now) {
      summary.overdueTasks++;
      detailedTasks.push({
        taskId: taskId,
        taskName: row[COL.TASK_NAME] || taskId,
        projectId: projectId,
        projectName: projectName,
        ownerName: ownerName,
        team: team,
        status: status,
        overdueHrs: Math.round((now - deadline) / 3600000),
        isStuck: true
      });
    }
    if (completedAt && completedAt >= weekStart) summary.tasksCompletedThisWeek++;
    if (isOverdue) summary.overdueTasks++;
    if (isActive) summary.activeTasks++;

    if (isDone && totalHrs > 0) {
      totalCompletionHours += totalHrs;
      completedWithTAT++;
    }

    if (createdAt) bumpTrend(createdAt.toISOString().substring(0, 10), 'created');
    if (completedAt) {
      var dStr = completedAt.toISOString().substring(0, 10);
      bumpTrend(dStr, 'completed');
      if (totalHrs > 0) bumpTrend(dStr, 'tat', totalHrs);
    }

    // NEW: Status Aging & Bottleneck detection (Active tasks only)
    if (isActive) {
      if (!statusAging[status]) statusAging[status] = { totalHrs: 0, count: 0 };
      var lastAction = row[23] ? new Date(row[23]) : (createdAt || now);
      var ageHrs = (now - lastAction) / 3600000;
      statusAging[status].totalHrs += ageHrs;
      statusAging[status].count++;

      // NEW: SLA Risk Detection (> 80% SLA consumed)
      var sla = slaConfig[row[2]] || { slaHours: 24 };
      if (ageHrs > (sla.slaHours * 0.8)) {
        teamSlaRisk[team] = (teamSlaRisk[team] || 0) + 1;
      }
    }

    byStatus[status] = (byStatus[status] || 0) + 1;
    var type = row[COL.TASK_TYPE] || 'General';
    byType[type] = (byType[type] || 0) + 1;
    var priority = row[COL.PRIORITY] || 'Medium';
    byPriority[priority] = (byPriority[priority] || 0) + 1;

    var u = ub(ownerEmail, ownerName, team);
    if (!u.typeStats[type]) u.typeStats[type] = { completed: 0, totalTAT: 0, active: 0, overdue: 0 };
    if (isDone) {
      u.completed++; u.totalTAT += totalHrs; if (breached) u.breaches++;
      u.typeStats[type].completed++;
      if (totalHrs > 0) u.typeStats[type].totalTAT += totalHrs;
    }
    else {
      u.active++; if (isOverdue) u.overdue++;
      u.typeStats[type].active++;
      if (isOverdue) u.typeStats[type].overdue++;
    }

    var t2 = tb(team);
    if (!t2.typeStats[type]) t2.typeStats[type] = { completed: 0, totalTAT: 0, active: 0, overdue: 0 };
    if (isDone) {
      t2.completed++;
      if (breached) t2.breaches++;
      if (totalHrs > 0) { t2.totalTAT += totalHrs; t2.tatCount++; }
      t2.typeStats[type].completed++;
      if (totalHrs > 0) t2.typeStats[type].totalTAT += totalHrs;
    } else {
      t2.active++;
      if (isOverdue) t2.overdue++;
      t2.typeStats[type].active++;
      if (isOverdue) t2.typeStats[type].overdue++;
    }

    var p2 = pb(projectId, projectName, team);
    p2.total++;
    if (isDone) p2.completed++;
    else p2.active++;
    if (isOverdue) p2.overdue++;
    if (breached) p2.breaches++;

    detailedTasks.push({
      taskId: row[COL.TASK_ID] || '',
      taskName: row[COL.TASK_NAME] || '',
      ownerEmail: ownerEmail,
      ownerName: ownerName,
      team: team,
      status: status,
      projectId: projectId || '',
      projectName: projectName,
      createdAt: createdAt ? createdAt.toISOString() : '',
      completedAt: completedAt ? completedAt.toISOString() : '',
      deadline: deadline ? deadline.toISOString() : '',
      totalHours: totalHrs,
      priority: priority,
      slaBreached: breached,
      isActive: isActive,
      isDone: isDone,
      isArchived: isArchived,
      isOverdue: !!isOverdue
    });
  }

  summary.averageCompletionTime = completedWithTAT > 0 ? parseFloat((totalCompletionHours / completedWithTAT).toFixed(1)) : 0;
  globalSummary.averageCompletionTime = globalCompletedWithTAT > 0 ? parseFloat((globalTotalCompletionHours / globalCompletedWithTAT).toFixed(1)) : 0;

  // Identify Global Bottleneck Team (using globalByTeam)
  var maxOverdueG = -1;
  var bottleneckG = 'None';
  Object.keys(globalByTeam).forEach(function (t) {
    if (globalByTeam[t].overdueTasks > maxOverdueG) {
      maxOverdueG = globalByTeam[t].overdueTasks;
      bottleneckG = t;
    }
  });

  var projectHandoffs = getProjectHandoffDetails_(filters, allTaskMeta, projectNameMap, memberMap);

  // Link routing counts to detailed tasks
  var routeMap = {};
  if (projectHandoffs && projectHandoffs.taskStats) {
    projectHandoffs.taskStats.forEach(function (ts) { routeMap[ts.taskId] = ts.routeCount; });
  }
  detailedTasks.forEach(function (dt) {
    if (dt.taskId && routeMap[dt.taskId]) dt.routeCount = routeMap[dt.taskId];
    else dt.routeCount = 0;
  });

  return {
    summary: summary, // Filtered
    globalSummary: globalSummary, // unfiltered
    globalBottleneck: bottleneckG,
    globalTeamStats: Object.keys(globalByTeam).map(function (t) { return { team: t, overdueTasks: globalByTeam[t].overdueTasks, avgTAT: globalByTeam[t].count > 0 ? (globalByTeam[t].totalTAT / globalByTeam[t].count) : 0 }; }),
    byUser: byUser,
    byTeam: byTeam,
    byStatus: byStatus,
    byType: byType,
    byPriority: byPriority,
    byProject: byProject,
    trendMap: trendMap,
    memberMap: memberMap,
    projectNameMap: projectNameMap,
    allTaskMeta: allTaskMeta,
    statusAging: statusAging,
    teamSlaRisk: teamSlaRisk,
    teamCapacity: teamCapacity,
    projectHandoffs: projectHandoffs,
    detailedTasks: detailedTasks
  };
}

// ------------------------------------------------------------------
// getDashboardSummary_
// ------------------------------------------------------------------
function getDashboardSummary_(agg) {
  return agg.summary;
}

// ------------------------------------------------------------------
// getPriorityBreakdown_
// Returns priority counts in a stable visual order
// ------------------------------------------------------------------
function getPriorityBreakdown_(agg) {
  var order = ['Critical', 'High', 'Medium', 'Low'];
  var labels = order.filter(function (k) { return (agg.byPriority[k] || 0) > 0; });
  Object.keys(agg.byPriority || {}).forEach(function (k) {
    if (order.indexOf(k) === -1 && (agg.byPriority[k] || 0) > 0) labels.push(k);
  });
  return {
    labels: labels,
    data: labels.map(function (k) { return agg.byPriority[k] || 0; })
  };
}


function getProjectOverview_(agg) {
  var projects = Object.keys(agg.byProject || {}).map(function (k) { return agg.byProject[k]; });
  var activeProjects = projects.filter(function (p) { return p.projectId !== 'NO_PROJECT' && p.active > 0; });
  var completedProjects = projects.filter(function (p) { return p.projectId !== 'NO_PROJECT' && p.total > 0 && p.active === 0 && p.completed > 0; });
  var atRiskProjects = projects.filter(function (p) { return p.projectId !== 'NO_PROJECT' && (p.overdue > 0 || p.breaches > 0); });
  var withOverdue = projects.filter(function (p) { return p.projectId !== 'NO_PROJECT' && p.overdue > 0; });
  var progressValues = projects.filter(function (p) { return p.projectId !== 'NO_PROJECT' && p.total > 0; }).map(function (p) { return Math.round((p.completed / p.total) * 100); });
  var avgProgress = progressValues.length ? Math.round(progressValues.reduce(function (s, v) { return s + v; }, 0) / progressValues.length) : 0;

  var topProjects = projects.slice().sort(function (a, b) { return b.total - a.total; }).slice(0, 8);
  var tasksByProject = {
    labels: topProjects.map(function (p) { return p.projectName; }),
    data: topProjects.map(function (p) { return p.total; })
  };

  var teamMap = {};
  projects.forEach(function (p) {
    var team = p.team || 'Unknown';
    if (!teamMap[team]) teamMap[team] = { ok: 0, risk: 0 };
    if (p.projectId === 'NO_PROJECT') return;
    if (p.overdue > 0 || p.breaches > 0) teamMap[team].risk++;
    else if (p.total > 0) teamMap[team].ok++;
  });
  var teamLabels = Object.keys(teamMap);

  return {
    kpis: {
      activeProjects: activeProjects.length,
      projectsAtRisk: atRiskProjects.length,
      completedProjects: completedProjects.length,
      noProjectTasks: (agg.byProject.NO_PROJECT && agg.byProject.NO_PROJECT.total) || 0,
      projectsWithOverdueTasks: withOverdue.length,
      avgProjectProgress: avgProgress
    },
    tasksByProject: tasksByProject,
    projectHealthByTeam: {
      labels: teamLabels,
      ok: teamLabels.map(function (t) { return teamMap[t].ok; }),
      risk: teamLabels.map(function (t) { return teamMap[t].risk; })
    },
    topProjects: topProjects.map(function (p) {
      return {
        projectId: p.projectId,
        projectName: p.projectName,
        total: p.total,
        active: p.active,
        completed: p.completed,
        overdue: p.overdue,
        progress: p.total ? Math.round((p.completed / p.total) * 100) : 0
      };
    })
  };
}

function getTeamKpis_(agg) {
  var teams = getTeamPerformanceStats_(agg);
  var onTrack = teams.filter(function (t) { return t.overdueTasks === 0 && t.breachRate < 20; }).length;
  var atRisk = teams.filter(function (t) { return t.overdueTasks > 0 || t.breachRate >= 20; }).length;
  var highestLoad = teams.length ? teams.slice().sort(function (a, b) { return b.activeTasks - a.activeTasks; })[0] : null;
  var tatTeams = teams.filter(function (t) { return t.avgTAT > 0; });
  var avgTat = tatTeams.length ? (tatTeams.reduce(function (s, t) { return s + t.avgTAT; }, 0) / tatTeams.length) : 0;
  return {
    teamsOnTrack: onTrack,
    teamsAtRisk: atRisk,
    highestActiveLoad: highestLoad ? highestLoad.activeTasks : 0,
    highestLoadTeam: highestLoad ? highestLoad.team : '',
    avgTeamTAT: Number(avgTat.toFixed(1))
  };
}

function getUserTaskReport_(agg) {
  var out = {};
  (agg.detailedTasks || []).forEach(function (t) {
    var key = String(t.ownerEmail || '').toLowerCase();
    if (!out[key]) out[key] = [];
    out[key].push(t);
  });
  Object.keys(out).forEach(function (k) {
    out[k].sort(function (a, b) {
      var da = new Date(a.completedAt || a.createdAt || 0).getTime();
      var db = new Date(b.completedAt || b.createdAt || 0).getTime();
      return db - da;
    });
  });
  return out;
}

// ------------------------------------------------------------------
// getUserPerformanceStats_
// Returns array sorted by completed desc
// ------------------------------------------------------------------
function getUserPerformanceStats_(agg) {
  return Object.keys(agg.byUser).map(function (email) {
    var u = agg.byUser[email];
    var total = u.completed + u.active;
    var rate = total > 0 ? Math.round((u.completed / total) * 100) : 0;
    var avgTat = u.completed > 0 ? parseFloat((u.totalTAT / u.completed).toFixed(1)) : 0;

    // Qualitative Labeling logic
    var label = 'Normal';
    if (avgTat > 0 && avgTat < 24 && u.active > 5) label = 'Fast response, high load';
    else if (avgTat > 72) label = 'Slow response, needs attention';
    else if (avgTat > 0 && avgTat < 24 && u.active <= 2) label = 'Very fast, low workload';
    else if (u.active > 8) label = 'High capacity, steady';

    return {
      email: email,
      name: u.name,
      team: u.team,
      activeTasks: u.active,
      completedTasks: u.completed,
      overdueTasks: u.overdue,
      slaBreaches: u.breaches,
      completionRate: rate,
      avgCompletionHours: avgTat,
      qualLabel: label,
      insights: generateUserInsights_(u, agg)
    };
  }).sort(function (a, b) { return b.completedTasks - a.completedTasks; });
}

function generateUserInsights_(u, agg) {
  var insights = [];
  var avgTat = u.completed > 0 ? (u.totalTAT / u.completed) : 0;

  if (u.overdue > u.active * 0.5) insights.push("User has critical backlog (50%+ tasks overdue).");
  if (avgTat > 72) insights.push("Response time is significantly higher than system average.");

  // Task-type specific insights
  Object.keys(u.typeStats || {}).forEach(function (type) {
    var stats = u.typeStats[type];
    if (stats.completed > 2) {
      var typeAvg = stats.totalTAT / stats.completed;
      if (typeAvg > avgTat * 1.5) insights.push("Response time is slower in " + type + " tasks.");
      if (typeAvg < avgTat * 0.5) insights.push("Exceptional speed in " + type + " tasks.");
    }
  });

  return insights;
}

// ------------------------------------------------------------------
// getTeamPerformanceStats_
// Returns array sorted by completion rate desc
// ------------------------------------------------------------------
function getTeamPerformanceStats_(agg) {
  return Object.keys(agg.byTeam).map(function (team) {
    var t = agg.byTeam[team];
    var total = t.completed + t.active;
    var rate = total > 0 ? Math.round((t.completed / total) * 100) : 0;
    return {
      team: team,
      activeTasks: t.active,
      completedTasks: t.completed,
      overdueTasks: t.overdue,
      completionRate: rate,
      avgTAT: t.tatCount > 0 ? parseFloat((t.totalTAT / t.tatCount).toFixed(1)) : 0,
      breachRate: t.completed > 0 ? Math.round((t.breaches / t.completed) * 100) : 0
    };
  }).sort(function (a, b) { return b.completionRate - a.completionRate; });
}

// ------------------------------------------------------------------
// getWorkloadDistribution_
// Returns user list with active task count + percentage of total
// ------------------------------------------------------------------
function getWorkloadDistribution_(agg) {
  var users = Object.keys(agg.byUser).map(function (email) {
    return { name: agg.byUser[email].name, email: email, team: agg.byUser[email].team, activeTasks: agg.byUser[email].active };
  }).filter(function (u) { return u.activeTasks > 0; });
  users.sort(function (a, b) { return b.activeTasks - a.activeTasks; });
  var total = users.reduce(function (s, u) { return s + u.activeTasks; }, 0);
  users.forEach(function (u) { u.pct = total > 0 ? Math.round((u.activeTasks / total) * 100) : 0; });
  return users;
}

// ------------------------------------------------------------------
// getTaskTrendData_
// Returns last 14 days of created vs completed counts
// ------------------------------------------------------------------
function getTaskTrendData_(agg) {
  var days = 14;
  var labels = [], created = [], completed = [];
  var now = new Date();
  for (var d = days - 1; d >= 0; d--) {
    var dt = new Date(now);
    dt.setDate(dt.getDate() - d);
    var ds = dt.toISOString().substring(0, 10);
    labels.push(ds.substring(5)); // MM-DD
    var bucket = agg.trendMap[ds] || {};
    created.push(bucket.created || 0);
    completed.push(bucket.completed || 0);
  }
  return { labels: labels, created: created, completed: completed };
}

// ------------------------------------------------------------------
// getGoalProgressStats_
// Reads Goals sheet + computes actual from task metrics in agg
// ------------------------------------------------------------------
function getGoalProgressStats_() {
  try {
    var sheet = getSheet(SHEETS.GOALS);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var taskSheet = getSheet(SHEETS.TASKS);
    var allTasks = taskSheet.getDataRange().getValues();
    var now = new Date();
    var out = [];

    // Read current dashboard filters to potentially hide goals not relevant to the current view
    // (Optional: currently we show all goals but compute progress based on the goal's own scope)

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var goalId = row[0];
      var goalName = row[1] || '';
      var scope = (row[2] || '').toLowerCase().trim(); // m:email or t:team or all
      var target = Number(row[3]) || 0;
      var metricType = row[4] || 'tasksCompleted';
      var startDate = row[5] ? new Date(row[5]) : null;
      var endDate = row[6] ? new Date(row[6]) : null;
      var desc = row[7] || '';

      var actual = 0;
      var gStart = startDate ? new Date(startDate).getTime() : 0;
      var gEnd = endDate ? new Date(endDate).getTime() : Infinity;

      // Parse scope
      var scopeType = 'all';
      var scopeID = '';
      if (scope.indexOf('t:') === 0) { scopeType = 'team'; scopeID = scope.substring(2); }
      else if (scope.indexOf('m:') === 0) { scopeType = 'member'; scopeID = scope.substring(2); }

      for (var j = 1; j < allTasks.length; j++) {
        var tRow = allTasks[j];
        if (isRecurringRow_(tRow)) continue;

        var tOwner = (tRow[3] || '').toLowerCase().trim();
        var tTeam = (tRow[5] || '').toLowerCase().trim();
        var tStatus = tRow[7];
        var tCompRaw = tRow[12]; // COMPLETED_AT
        if (!tCompRaw) continue;

        var tDate = new Date(tCompRaw).getTime();
        if (isNaN(tDate)) continue;
        if (tDate < gStart || tDate > gEnd) continue;

        var isDone = tStatus === 'Done' || tStatus === 'Archived' || tStatus === 'Completed' || tRow[7] === 'Done';
        if (!isDone) continue;

        if (metricType === 'tasksClosed' && tStatus !== 'Archived') continue;

        // Apply Scope Filter
        if (scopeType === 'member') {
          if (tOwner !== scopeID) continue;
        } else if (scopeType === 'team') {
          if (tTeam !== scopeID) continue;
        }

        actual++;
      }

      var pct = target > 0 ? Math.min(100, Math.round((actual / target) * 100)) : 0;
      var daysLeft = endDate ? Math.ceil((endDate - now) / 86400000) : null;
      var status = pct >= 100 ? 'Achieved' : (daysLeft !== null && daysLeft < 0) ? 'Overdue' : 'In Progress';

      out.push({
        goalId: goalId, goalName: goalName, ownerEmail: scope, target: target, actual: actual, pct: pct,
        metricType: metricType, daysLeft: daysLeft, status: status, description: desc,
        startDate: startDate ? startDate.toISOString() : '', endDate: endDate ? endDate.toISOString() : ''
      });
    }
    return out;
  } catch (e) {
    console.warn('getGoalProgressStats_: ' + e.message);
    return [];
  }
}

// ------------------------------------------------------------------
// getAnalyticsDashboard — MAIN API ENDPOINT
// Called by: Dashboard tabs (Goals, Team, User, Tasks)
// Returns structured payload cached for 5 minutes per filter set
// ------------------------------------------------------------------
function getAnalyticsDashboard(filters) {
  try {
    var actor = requireRole_(['Owner', 'Manager']);
    filters = filters || {};

    // Manager scoping: force team filter
    if (actor.role === 'Manager') {
      if (!actor.team) return err_('UNAUTHORIZED');
      filters.team = actor.team;
    }

    // Try cache first
    var cacheKey = dashCacheKey_(filters);
    var cache = CacheService.getScriptCache();
    try {
      var hit = cache.get(cacheKey);
      if (hit) return ok_(JSON.parse(hit));
    } catch (e) { /* cache miss is fine */ }

    // Single aggregation pass — all metrics share this
    var agg = aggregateTasks_(filters);

    var projectOverview = getProjectOverview_(agg);
    var teamPerformance = getTeamPerformanceStats_(agg);
    var payload = {
      summary: getDashboardSummary_(agg),
      globalSummary: agg.globalSummary,
      globalBottleneck: agg.globalBottleneck,
      globalTeamStats: agg.globalTeamStats,
      projectOverview: projectOverview,
      tasksByStatus: { labels: Object.keys(agg.byStatus), data: Object.values(agg.byStatus) },
      tasksByType: { labels: Object.keys(agg.byType), data: Object.values(agg.byType) },
      tasksByPriority: getPriorityBreakdown_(agg),
      tasksByProject: projectOverview.tasksByProject,
      tatByTeam: {
        labels: Object.keys(agg.byTeam).filter(function (k) { return agg.byTeam[k].tatCount > 0; }),
        data: Object.keys(agg.byTeam).filter(function (k) { return agg.byTeam[k].tatCount > 0; }).map(function (k) { return Number((agg.byTeam[k].totalTAT / agg.byTeam[k].tatCount).toFixed(1)); })
      },
      userPerformance: getUserPerformanceStats_(agg),
      userTaskReport: getUserTaskReport_(agg),
      teamPerformance: teamPerformance,
      teamKpis: getTeamKpis_(agg),
      goalProgress: getGoalProgressStats_(),
      workloadDistribution: getWorkloadDistribution_(agg),
      taskTrend: getTaskTrendData_(agg),
      stuckTasks: (agg.detailedTasks || []).filter(function (t) { return t.isStuck; }).sort(function (a, b) { return b.overdueHrs - a.overdueHrs; }).slice(0, 5),
      statusAging: agg.statusAging,
      teamSlaRisk: agg.teamSlaRisk,
      teamCapacity: agg.teamCapacity,
      projectHandoffs: agg.projectHandoffs
    };

    // Cache for 5 minutes — heavy enough to be worth it
    try { cache.put(cacheKey, JSON.stringify(payload), 300); } catch (e) { }

    return ok_(payload);

  } catch (e) {
    if (e.message === 'UNAUTHORIZED') return err_('UNAUTHORIZED');
    console.error('getAnalyticsDashboard: ' + e.message);
    return err_('SYSTEM_ERROR');
  }
}

// ------------------------------------------------------------------
// invalidateDashCache_ — call after any task create/complete/route
// Keeps dashboard fresh without waiting for 5-minute TTL
// ------------------------------------------------------------------
function invalidateDashCache_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var v = parseInt(props.getProperty('DASH_CACHE_VERSION') || '1') + 1;
    props.setProperty('DASH_CACHE_VERSION', v.toString());
    console.log('invalidateDashCache_: version bumped to ' + v);
  } catch (e) {
    console.warn('invalidateDashCache_ failed: ' + e.message);
  }
}

/**
 * getProjectHandoffDetails_ — Sprint 10 Strategic Analysis
 * Analyzes the EventLog to trace task movement and holding time per person.
 */
function getProjectHandoffDetails_(filters, taskMetaMap, projectNameMap, memberMap) {
  var projectId = filters.project || '';
  var teamFilter = filters.team || '';
  var memberFilter = String(filters.member || '').toLowerCase();
  var fromDate = filters.from ? new Date(filters.from) : getDateDaysAgo_(30);
  var toDate = filters.to ? new Date(filters.to) : new Date();
  taskMetaMap = taskMetaMap || {};
  projectNameMap = projectNameMap || {};
  memberMap = memberMap || {};
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);
  try {
    var logs = getSheet(SHEETS.EVENT_LOG).getDataRange().getValues();
    var now = new Date();

    var totalWaitHrs = 0;
    var waitCount = 0;
    var crossTeamRoutes = 0;
    var sameTeamRoutes = 0;
    var participants = {};   // email -> { email, name, totalTime, actions, breaches }
    var handoffs = [];       // Array of routed events
    var lifecycleMap = {};   // taskId -> [{ owner, durationHrs }]
    var lifecycleEvents = {}; // taskId -> [{ owner, ts, final }]
    var projectStats = {};   // projectId -> aggregate
    var taskStats = {};      // taskId -> aggregate
    var teamPairs = {};      // from->to aggregate

    function metaForTask_(taskId) {
      return taskMetaMap[taskId] || {};
    }

    function ensureProjectStat_(projId, projName) {
      var key = projId || 'NO_PROJECT';
      if (!projectStats[key]) {
        projectStats[key] = {
          projectId: key,
          projectName: projName || 'No Project',
          routeCount: 0,
          totalWaitHrs: 0,
          crossTeamRoutes: 0,
          sameTeamRoutes: 0,
          taskIds: {}
        };
      }
      return projectStats[key];
    }

    function ensureTaskStat_(taskId, meta, projId, projName) {
      if (!taskStats[taskId]) {
        taskStats[taskId] = {
          taskId: taskId,
          taskName: meta.taskName || taskId,
          projectId: projId || meta.projectId || '',
          projectName: projName || meta.projectName || 'No Project',
          routeCount: 0,
          totalWaitHrs: 0,
          lastRoutedAt: ''
        };
      }
      return taskStats[taskId];
    }

    function ensureLifecycle_(taskId) {
      if (!lifecycleEvents[taskId]) lifecycleEvents[taskId] = [];
      return lifecycleEvents[taskId];
    }

    for (var i = 1; i < logs.length; i++) {
      var row = logs[i];
      if (!row[1] || !row[2]) continue; // Skip malformed rows

      var type = row[1];
      var taskId = row[2];
      var meta = metaForTask_(taskId);
      var projId = row[3] || meta.projectId || '';
      var projName = projId ? (projectNameMap[projId] || meta.projectName || projId) : 'No Project';
      var actor = String(row[4] || '').toLowerCase();
      var actorName = (memberMap[actor] && memberMap[actor].name) || actor;
      var actorTeam = row[5] || meta.team || '';
      var target = String(row[6] || '').toLowerCase();
      var targetTeam = row[7] || meta.team || '';
      var timeSpent = Number(row[10]) || 0;
      var breached = row[12] === true;
      var ts = row[14] ? new Date(row[14]) : null;
      if (!ts || isNaN(ts.getTime())) ts = now;
      if (ts < fromDate || ts > toDate) continue;
      if (projectId && projId !== projectId) continue;
      if (teamFilter && actorTeam !== teamFilter && targetTeam !== teamFilter && meta.team !== teamFilter) continue;
      if (memberFilter && actor !== memberFilter && target !== memberFilter && String(meta.ownerEmail || '').toLowerCase() !== memberFilter) continue;

      if (type === EVENT.TASK_CREATED) {
        var createdOwner = target || String(meta.ownerEmail || '').toLowerCase() || actor;
        if (createdOwner) ensureLifecycle_(taskId).push({ owner: createdOwner, ts: ts.getTime(), final: false });
        continue;
      }

      if (type === EVENT.TASK_REOPENED) {
        var reopenedOwner = target || String(meta.ownerEmail || '').toLowerCase() || actor;
        if (reopenedOwner) ensureLifecycle_(taskId).push({ owner: reopenedOwner, ts: ts.getTime(), final: false });
        continue;
      }

      if (type === EVENT.TASK_ROUTED && target) {
        totalWaitHrs += timeSpent;
        waitCount++;
        if (actor) {
          if (!participants[actor]) participants[actor] = { email: actor, name: actorName, totalTime: 0, actions: 0, breaches: 0 };
          participants[actor].totalTime += timeSpent;
          participants[actor].actions++;
          if (breached) participants[actor].breaches++;
        }

        var isCrossTeam = !!(actorTeam && targetTeam && actorTeam !== targetTeam);
        if (isCrossTeam) crossTeamRoutes++;
        else sameTeamRoutes++;

        var pStat = ensureProjectStat_(projId, projName);
        pStat.routeCount++;
        pStat.totalWaitHrs += timeSpent;
        pStat.taskIds[taskId] = true;
        if (isCrossTeam) pStat.crossTeamRoutes++;
        else pStat.sameTeamRoutes++;

        var tStat = ensureTaskStat_(taskId, meta, projId, projName);
        tStat.routeCount++;
        tStat.totalWaitHrs += timeSpent;
        tStat.lastRoutedAt = ts.toISOString();

        var pairLabel = (actorTeam || 'Unknown') + ' -> ' + (targetTeam || 'Unknown');
        if (!teamPairs[pairLabel]) {
          teamPairs[pairLabel] = { label: pairLabel, fromTeam: actorTeam || 'Unknown', toTeam: targetTeam || 'Unknown', routeCount: 0, totalWaitHrs: 0 };
        }
        teamPairs[pairLabel].routeCount++;
        teamPairs[pairLabel].totalWaitHrs += timeSpent;

        handoffs.push({
          taskId: taskId,
          taskName: meta.taskName || taskId,
          projectId: projId,
          projectName: projName,
          from: actor,
          fromTeam: actorTeam,
          to: target,
          toTeam: targetTeam,
          waitHrs: timeSpent,
          timestamp: ts.toISOString()
        });
        ensureLifecycle_(taskId).push({ owner: target, ts: ts.getTime(), final: false });
        continue;
      }

      if (type === EVENT.TASK_COMPLETED || type === EVENT.TASK_ARCHIVED) {
        var finalOwner = actor || target || String(meta.ownerEmail || '').toLowerCase();
        if (finalOwner) ensureLifecycle_(taskId).push({ owner: finalOwner, ts: ts.getTime(), final: true });
      }
    }

    Object.keys(taskStats).forEach(function (tid) {
      var events = (lifecycleEvents[tid] || []).slice().sort(function (a, b) { return a.ts - b.ts; });
      if (!events.length) return;
      var segments = [];
      for (var j = 0; j < events.length; j++) {
        var ev = events[j];
        if (!ev.owner) continue;
        var nextTs = (j < events.length - 1) ? events[j + 1].ts : (ev.final ? ev.ts : now.getTime());
        var dur = ev.final ? 0.1 : Math.max(0.1, (nextTs - ev.ts) / 3600000);
        if (!isFinite(dur)) dur = 0.1;
        dur = Number(dur.toFixed(2));
        var prev = segments.length ? segments[segments.length - 1] : null;
        if (prev && prev.owner === ev.owner) prev.durationHrs = Number((prev.durationHrs + dur).toFixed(2));
        else segments.push({ owner: ev.owner, durationHrs: dur });
      }
      if (segments.length) lifecycleMap[tid] = segments;
    });

    var participantList = Object.keys(participants).map(function (email) {
      var p = participants[email];
      p.avgWaitHrs = p.actions > 0 ? Number((p.totalTime / p.actions).toFixed(2)) : 0;
      return p;
    }).sort(function (a, b) { return b.actions - a.actions; });

    var projectList = Object.keys(projectStats).map(function (key) {
      var p = projectStats[key];
      var uniqueTasks = Object.keys(p.taskIds).length;
      return {
        projectId: p.projectId,
        projectName: p.projectName,
        routeCount: p.routeCount,
        uniqueTasks: uniqueTasks,
        totalWaitHrs: Number(p.totalWaitHrs.toFixed(2)),
        avgWaitHrs: p.routeCount > 0 ? Number((p.totalWaitHrs / p.routeCount).toFixed(2)) : 0,
        crossTeamRoutes: p.crossTeamRoutes,
        sameTeamRoutes: p.sameTeamRoutes,
        avgRoutesPerTask: uniqueTasks > 0 ? Number((p.routeCount / uniqueTasks).toFixed(2)) : 0
      };
    }).sort(function (a, b) { return b.routeCount - a.routeCount; });

    var taskList = Object.keys(taskStats).map(function (key) {
      var t = taskStats[key];
      return {
        taskId: t.taskId,
        taskName: t.taskName,
        projectId: t.projectId,
        projectName: t.projectName,
        routeCount: t.routeCount,
        totalWaitHrs: Number(t.totalWaitHrs.toFixed(2)),
        avgWaitHrs: t.routeCount > 0 ? Number((t.totalWaitHrs / t.routeCount).toFixed(2)) : 0,
        lastRoutedAt: t.lastRoutedAt
      };
    }).sort(function (a, b) {
      if (b.routeCount !== a.routeCount) return b.routeCount - a.routeCount;
      return b.totalWaitHrs - a.totalWaitHrs;
    });

    var teamPairList = Object.keys(teamPairs).map(function (key) {
      var pair = teamPairs[key];
      return {
        label: pair.label,
        fromTeam: pair.fromTeam,
        toTeam: pair.toTeam,
        routeCount: pair.routeCount,
        avgWaitHrs: pair.routeCount > 0 ? Number((pair.totalWaitHrs / pair.routeCount).toFixed(2)) : 0
      };
    }).sort(function (a, b) { return b.routeCount - a.routeCount; });

    return {
      handoffs: handoffs.sort(function (a, b) { return new Date(b.timestamp) - new Date(a.timestamp); }).slice(0, 200),
      participants: participantList,
      lifecycleMap: lifecycleMap,
      averageHandoffWait: waitCount > 0 ? Number((totalWaitHrs / waitCount).toFixed(2)) : 0,
      projectStats: projectList,
      taskStats: taskList,
      teamPairs: teamPairList,
      summary: {
        totalRoutes: handoffs.length,
        routedTasks: taskList.length,
        projectsCovered: projectList.length,
        crossTeamRoutes: crossTeamRoutes,
        sameTeamRoutes: sameTeamRoutes,
        crossTeamPct: handoffs.length ? Math.round((crossTeamRoutes / handoffs.length) * 100) : 0,
        churnyTasks: taskList.filter(function (t) { return t.routeCount >= 4; }).length
      },
      selectedProjectId: projectId || '',
      selectedProjectName: projectId ? (projectNameMap[projectId] || projectId) : ''
    };
  } catch (e) {
    console.error('getProjectHandoffDetails_ failed: ' + e.message);
    return { handoffs: [], participants: [], lifecycleMap: {}, projectStats: [], taskStats: [], teamPairs: [], summary: {} };
  }
}
// Called on a 4-minute time-based trigger so the 5-minute cache is
// always warm. First load after cache expiry becomes a cache hit.
// Run installTriggers() to register the trigger.
//
// Strategy: pre-build the all-teams no-filter payload (Owner view),
// which is the heaviest query and the one most likely to be stale.
// Manager-scoped views are lighter and warm themselves on first access.
// ------------------------------------------------------------------
function warmDashCache() {
  try {
    var now = new Date();
    var thirtyAgo = new Date(now.getTime() - 30 * 24 * 3600000);
    var filters = {
      team: '',
      member: '',
      from: thirtyAgo.toISOString().substring(0, 10),
      to: now.toISOString().substring(0, 10)
    };

    var cacheKey = dashCacheKey_(filters);
    var cache = CacheService.getScriptCache();

    // Only recompute if cache is already expired or nearly expired (< 60s left)
    try {
      var existing = cache.get(cacheKey);
      if (existing) {
        console.log('warmDashCache: cache still warm, skipping recompute');
        return;
      }
    } catch (e) { /* miss is expected — proceed */ }

    // Recompute and store — same logic as getAnalyticsDashboard without auth guard
    var agg = aggregateTasks_(filters);
    var projectOverview = getProjectOverview_(agg);
    var teamPerformance = getTeamPerformanceStats_(agg);
    var payload = {
      summary: getDashboardSummary_(agg),
      projectOverview: projectOverview,
      tasksByStatus: { labels: Object.keys(agg.byStatus), data: Object.values(agg.byStatus) },
      tasksByType: { labels: Object.keys(agg.byType), data: Object.values(agg.byType) },
      tasksByPriority: getPriorityBreakdown_(agg),
      tasksByProject: projectOverview.tasksByProject,
      tatByTeam: {
        labels: Object.keys(agg.byTeam).filter(function (k) { return agg.byTeam[k].tatCount > 0; }),
        data: Object.keys(agg.byTeam).filter(function (k) { return agg.byTeam[k].tatCount > 0; })
          .map(function (k) { return Number((agg.byTeam[k].totalTAT / agg.byTeam[k].tatCount).toFixed(1)); })
      },
      userPerformance: getUserPerformanceStats_(agg),
      userTaskReport: getUserTaskReport_(agg),
      teamPerformance: teamPerformance,
      teamKpis: getTeamKpis_(agg),
      goalProgress: getGoalProgressStats_(),
      workloadDistribution: getWorkloadDistribution_(agg),
      taskTrend: getTaskTrendData_(agg),
      statusAging: agg.statusAging,
      teamSlaRisk: agg.teamSlaRisk,
      teamCapacity: agg.teamCapacity,
      projectHandoffs: agg.projectHandoffs
    };

    try { cache.put(cacheKey, JSON.stringify(payload), 300); } catch (e) { }
    console.log('warmDashCache: recomputed at ' + now.toISOString());

  } catch (e) {
    // Non-fatal — warm cache is an optimisation, not a requirement
    console.warn('warmDashCache failed: ' + e.message);
  }
}
