const TIME_ZONE = 'Asia/Tokyo';
const MAX_ROWS = 1000;

function createProjectPortfolio() {
  assertSeedAvailable_();
  const today = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd');
  const ss = SpreadsheetApp.create('プロジェクト横断管理_' + today);
  ss.setSpreadsheetLocale('ja_JP');
  ss.setSpreadsheetTimeZone(TIME_ZONE);

  const dashboard = ss.getSheets()[0];
  dashboard.setName('Dashboard');
  const projects = ss.insertSheet('Projects');
  const tasks = ss.insertSheet('Tasks');
  const risks = ss.insertSheet('Risks & Milestones');
  const dependencies = ss.insertSheet('Dependencies');
  const milestoneTimeline = ss.insertSheet('Milestone Timeline');
  const decisionLog = ss.insertSheet('Decision Log');
  const master = ss.insertSheet('Master');

  buildMaster_(master);
  buildDependencies_(dependencies, projects, master, false);
  buildProjects_(projects, projectRowsFromSeed_(), master);
  buildTasks_(tasks, projects, master, false);
  buildRisks_(risks, projects, master, false);
  buildMilestoneTimeline_(milestoneTimeline);
  buildDecisionLog_(decisionLog, projects, master, false);
  buildDashboard_(dashboard, PROJECT_SEED_META);
  orderSheets_(ss);

  ss.setActiveSheet(dashboard);
  SpreadsheetApp.flush();
  const report = verifyWorkbook_(ss);
  console.log('CREATED_URL=' + ss.getUrl());
  console.log('VERIFICATION=' + JSON.stringify(report));
  return ss.getUrl();
}

function upgradeProjectPortfolio(spreadsheetId) {
  if (!spreadsheetId) throw new Error('spreadsheetIdを指定してください');
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const dashboard = getOrInsertSheet_(ss, 'Dashboard');
  const projects = getOrInsertSheet_(ss, 'Projects');
  const tasks = getOrInsertSheet_(ss, 'Tasks');
  const risks = getOrInsertSheet_(ss, 'Risks & Milestones');
  const dependencies = getOrInsertSheet_(ss, 'Dependencies');
  const milestoneTimeline = getOrInsertSheet_(ss, 'Milestone Timeline');
  const decisionLog = getOrInsertSheet_(ss, 'Decision Log');
  const master = getOrInsertSheet_(ss, 'Master');

  upgradeMaster_(master);
  buildDependencies_(dependencies, projects, master, true);
  upgradeProjects_(projects, master);
  buildTasks_(tasks, projects, master, true);
  buildRisks_(risks, projects, master, true);
  buildMilestoneTimeline_(milestoneTimeline);
  buildDecisionLog_(decisionLog, projects, master, true);

  const projectCount = projects.getRange('A2:A' + MAX_ROWS).getValues().filter(function(row) {
    return row[0] !== '';
  }).length;
  buildDashboard_(dashboard, {
    source: '既存Projects',
    projectCount: projectCount,
    generatedAt: Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd')
  });
  orderSheets_(ss);

  ss.setActiveSheet(dashboard);
  SpreadsheetApp.flush();
  const report = verifyWorkbook_(ss);
  console.log('UPGRADED_URL=' + ss.getUrl());
  console.log('VERIFICATION=' + JSON.stringify(report));
  return ss.getUrl();
}

function upgradeConfiguredProjectPortfolio() {
  if (typeof TARGET_SPREADSHEET_ID === 'undefined' || !TARGET_SPREADSHEET_ID) {
    throw new Error('非公開設定のTARGET_SPREADSHEET_IDが必要です');
  }
  return upgradeProjectPortfolio(TARGET_SPREADSHEET_ID);
}

function verifyProjectPortfolio(spreadsheetId) {
  const report = verifyWorkbook_(SpreadsheetApp.openById(spreadsheetId));
  console.log('VERIFICATION=' + JSON.stringify(report));
  return report;
}

function verifyConfiguredProjectPortfolio() {
  if (typeof TARGET_SPREADSHEET_ID === 'undefined' || !TARGET_SPREADSHEET_ID) {
    throw new Error('非公開設定のTARGET_SPREADSHEET_IDが必要です');
  }
  return verifyProjectPortfolio(TARGET_SPREADSHEET_ID);
}

function assertSeedAvailable_() {
  if (typeof PROJECT_SEED === 'undefined' || typeof PROJECT_SEED_META === 'undefined') {
    throw new Error('Project Seedをビルドしてから実行してください');
  }
}

function projectRowsFromSeed_() {
  return PROJECT_SEED.map(function(project) {
    return [
      project.projectId,
      project.category,
      project.name,
      project.localPath,
      project.managementType,
      project.status,
      project.priority,
      project.owner,
      project.nextAction,
      project.dueDate,
      project.repositoryUrl,
      project.memo,
      new Date(project.lastConfirmedAt + 'T00:00:00+09:00')
    ];
  });
}

function getOrInsertSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function orderSheets_(ss) {
  [
    'Dashboard', 'Projects', 'Tasks', 'Risks & Milestones', 'Dependencies',
    'Milestone Timeline', 'Decision Log', 'Master'
  ].forEach(function(name, index) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });
}
