const SINGLE_TIME_ZONE = 'Asia/Tokyo';
const SINGLE_MAX_ROWS = 1000;
const SINGLE_SHEET_NAMES = Object.freeze({
  dashboard: 'ダッシュボード', guide: 'ガイド', charter: 'プロジェクト憲章', workItems: '作業項目',
  raid: 'RAIDログ', stakeholders: 'ステークホルダー', decisions: '意思決定・変更ログ', master: 'マスター'
});
const SINGLE_SHEET_ORDER = Object.freeze([
  SINGLE_SHEET_NAMES.dashboard, SINGLE_SHEET_NAMES.guide, SINGLE_SHEET_NAMES.charter, SINGLE_SHEET_NAMES.workItems,
  SINGLE_SHEET_NAMES.raid, SINGLE_SHEET_NAMES.stakeholders, SINGLE_SHEET_NAMES.decisions, SINGLE_SHEET_NAMES.master
]);

function createSingleProjectWorkbook() {
  singleAssertSeed_();
  const ss = SpreadsheetApp.create('単一プロジェクト管理_' + Utilities.formatDate(new Date(), SINGLE_TIME_ZONE, 'yyyy-MM-dd'));
  ss.setSpreadsheetLocale('ja_JP');
  ss.setSpreadsheetTimeZone(SINGLE_TIME_ZONE);
  const dashboard = ss.getSheets()[0];
  dashboard.setName(SINGLE_SHEET_NAMES.dashboard);
  SINGLE_SHEET_ORDER.slice(1).forEach(function(name) { ss.insertSheet(name); });
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
  singleBuildMaster_(singleSheet_(ss, SINGLE_SHEET_NAMES.master));
  singleBuildGuide_(singleSheet_(ss, SINGLE_SHEET_NAMES.guide));
  singleBuildCharter_(singleSheet_(ss, SINGLE_SHEET_NAMES.charter));
  singleBuildWorkItems_(singleSheet_(ss, SINGLE_SHEET_NAMES.workItems));
  singleBuildRaid_(singleSheet_(ss, SINGLE_SHEET_NAMES.raid));
  singleBuildStakeholders_(singleSheet_(ss, SINGLE_SHEET_NAMES.stakeholders));
  singleBuildDecisions_(singleSheet_(ss, SINGLE_SHEET_NAMES.decisions));
  singleBuildDashboard_(singleSheet_(ss, SINGLE_SHEET_NAMES.dashboard));
  SINGLE_SHEET_ORDER.forEach(function(name, index) {
    ss.setActiveSheet(singleSheet_(ss, name)); ss.moveActiveSheet(index + 1);
  });
}
