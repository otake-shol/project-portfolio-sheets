const SINGLE_TIME_ZONE = 'Asia/Tokyo';
const SINGLE_MAX_ROWS = 1000;

function createSingleProjectWorkbook() {
  singleAssertSeed_();
  const ss = SpreadsheetApp.create('単一プロジェクト管理_' + Utilities.formatDate(new Date(), SINGLE_TIME_ZONE, 'yyyy-MM-dd'));
  ss.setSpreadsheetLocale('ja_JP');
  ss.setSpreadsheetTimeZone(SINGLE_TIME_ZONE);
  const dashboard = ss.getSheets()[0];
  dashboard.setName('Dashboard');
  ['Guide', 'Project Charter', 'Work Items', 'RAID Log', 'Stakeholders', 'Decision Log', 'Master'].forEach(function(name) { ss.insertSheet(name); });
  buildSingleWorkbook_(ss);
  ss.setActiveSheet(dashboard);
  SpreadsheetApp.flush();
  const report = verifySingleProjectWorkbook_(ss);
  console.log('CREATED_URL=' + ss.getUrl());
  console.log('VERIFICATION=' + JSON.stringify(report));
  return ss.getUrl();
}

function verifySingleProjectWorkbook(spreadsheetId) {
  if (!spreadsheetId) throw new Error('spreadsheetIdを指定してください');
  return verifySingleProjectWorkbook_(SpreadsheetApp.openById(spreadsheetId));
}

function singleAssertSeed_() {
  if (typeof SINGLE_PROJECT_SEED === 'undefined' || typeof SINGLE_PROJECT_SEED_META === 'undefined') throw new Error('Single Project Seedをビルドしてから実行してください');
}

function singleDate_(value) {
  return value ? new Date(value + 'T00:00:00+09:00') : '';
}

function singleSheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function buildSingleWorkbook_(ss) {
  singleBuildMaster_(singleSheet_(ss, 'Master'));
  singleBuildGuide_(singleSheet_(ss, 'Guide'));
  singleBuildCharter_(singleSheet_(ss, 'Project Charter'));
  singleBuildWorkItems_(singleSheet_(ss, 'Work Items'));
  singleBuildRaid_(singleSheet_(ss, 'RAID Log'));
  singleBuildStakeholders_(singleSheet_(ss, 'Stakeholders'));
  singleBuildDecisions_(singleSheet_(ss, 'Decision Log'));
  singleBuildDashboard_(singleSheet_(ss, 'Dashboard'));
  ['Dashboard', 'Guide', 'Project Charter', 'Work Items', 'RAID Log', 'Stakeholders', 'Decision Log', 'Master'].forEach(function(name, index) {
    ss.setActiveSheet(singleSheet_(ss, name)); ss.moveActiveSheet(index + 1);
  });
}
