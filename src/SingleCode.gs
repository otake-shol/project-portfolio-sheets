const SINGLE_TIME_ZONE = 'Asia/Tokyo';
const SINGLE_MAX_ROWS = 1000;
const SINGLE_GANTT_PERIODS = 52;
const SINGLE_SUMMARY_PERIODS = 18;
const SINGLE_SUMMARY_ITEMS = 18;
const SINGLE_SHEET_NAMES = Object.freeze({
  dashboard: 'ダッシュボード', guide: 'ガイド', charter: 'プロジェクト憲章', summarySchedule: '概要スケジュール', workItems: 'WBS', gantt: 'ガントチャート',
  raid: 'RAIDログ', stakeholders: 'ステークホルダー', decisions: '意思決定・変更ログ', master: 'マスター'
});
const SINGLE_SHEET_ORDER = Object.freeze([
  SINGLE_SHEET_NAMES.dashboard, SINGLE_SHEET_NAMES.guide, SINGLE_SHEET_NAMES.charter, SINGLE_SHEET_NAMES.summarySchedule, SINGLE_SHEET_NAMES.workItems,
  SINGLE_SHEET_NAMES.gantt, SINGLE_SHEET_NAMES.raid, SINGLE_SHEET_NAMES.stakeholders, SINGLE_SHEET_NAMES.decisions, SINGLE_SHEET_NAMES.master
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

function rebuildSingleProjectWorkbook(spreadsheetId) {
  if (!spreadsheetId) throw new Error('spreadsheetIdを指定してください');
  return rebuildSingleProjectWorkbook_(SpreadsheetApp.openById(spreadsheetId));
}

function rebuildActiveSingleProjectWorkbook() {
  return rebuildSingleProjectWorkbook_(SpreadsheetApp.getActiveSpreadsheet());
}

function rebuildSingleProjectWorkbook_(ss) {
  singleAssertSeed_();
  const legacyWorkItems = ss.getSheetByName('作業項目');
  if (legacyWorkItems && !ss.getSheetByName(SINGLE_SHEET_NAMES.workItems)) legacyWorkItems.setName(SINGLE_SHEET_NAMES.workItems);
  SINGLE_SHEET_ORDER.forEach(function(name) { singleSheet_(ss, name); });
  buildSingleWorkbook_(ss);
  ss.setActiveSheet(singleSheet_(ss, SINGLE_SHEET_NAMES.dashboard));
  SpreadsheetApp.flush();
  const report = verifySingleProjectWorkbook_(ss);
  console.log('UPDATED_URL=' + ss.getUrl());
  console.log('VERIFICATION=' + JSON.stringify(report));
  return report;
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
  singleBuildSummarySchedule_(singleSheet_(ss, SINGLE_SHEET_NAMES.summarySchedule));
  singleBuildGantt_(singleSheet_(ss, SINGLE_SHEET_NAMES.gantt));
  singleBuildRaid_(singleSheet_(ss, SINGLE_SHEET_NAMES.raid));
  singleBuildStakeholders_(singleSheet_(ss, SINGLE_SHEET_NAMES.stakeholders));
  singleBuildDecisions_(singleSheet_(ss, SINGLE_SHEET_NAMES.decisions));
  singleBuildDashboard_(singleSheet_(ss, SINGLE_SHEET_NAMES.dashboard));
  SINGLE_SHEET_ORDER.forEach(function(name, index) {
    ss.setActiveSheet(singleSheet_(ss, name)); ss.moveActiveSheet(index + 1);
  });
}
