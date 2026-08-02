function verifySingleProjectWorkbook_(ss) {
  SpreadsheetApp.flush();
  const expected = SINGLE_SHEET_ORDER;
  const names = ss.getSheets().map(function(sheet) { return sheet.getName(); });
  const missingSheets = expected.filter(function(name) { return names.indexOf(name) === -1; });
  const errors = [];
  ss.getSheets().forEach(function(sheet) { sheet.getDataRange().getDisplayValues().forEach(function(row, rowIndex) { row.forEach(function(value, columnIndex) { if (/^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|ERROR!)/.test(value)) errors.push(sheet.getName() + '!' + sheet.getRange(rowIndex + 1, columnIndex + 1).getA1Notation()); }); }); });
  const dashboard = ss.getSheetByName(SINGLE_SHEET_NAMES.dashboard); const work = ss.getSheetByName(SINGLE_SHEET_NAMES.workItems); const gantt = ss.getSheetByName(SINGLE_SHEET_NAMES.gantt); const raid = ss.getSheetByName(SINGLE_SHEET_NAMES.raid); const stakeholders = ss.getSheetByName(SINGLE_SHEET_NAMES.stakeholders); const decision = ss.getSheetByName(SINGLE_SHEET_NAMES.decisions);
  const report = { sheets: names, missingSheets: missingSheets, formulaErrors: errors, validations: {}, filters: {}, frozenHeaders: {}, conditionalFormats: {}, dashboardFormulas: [], ganttFormulas: [] };
  if (missingSheets.length) { report.passed = false; return report; }
  report.validations = { wbsLevel: Boolean(work.getRange('C2').getDataValidation()), wbsType: Boolean(work.getRange('D2').getDataValidation()), wbsStatus: Boolean(work.getRange('F2').getDataValidation()), ganttUnit: Boolean(gantt.getRange('E2').getDataValidation()), ganttPeriods: Boolean(gantt.getRange('H2').getDataValidation()), raidType: Boolean(raid.getRange('B2').getDataValidation()), decisionType: Boolean(decision.getRange('B2').getDataValidation()), decisionStatus: Boolean(decision.getRange('H2').getDataValidation()) };
  report.filters = { wbs: Boolean(work.getFilter()), raid: Boolean(raid.getFilter()), stakeholders: Boolean(stakeholders.getFilter()), decisions: Boolean(decision.getFilter()) };
  expected.forEach(function(name) { const sheet = ss.getSheetByName(name); report.frozenHeaders[name] = name === SINGLE_SHEET_NAMES.dashboard ? sheet.getFrozenRows() >= 2 : sheet.getFrozenRows() >= 1; });
  [SINGLE_SHEET_NAMES.dashboard, SINGLE_SHEET_NAMES.workItems, SINGLE_SHEET_NAMES.gantt, SINGLE_SHEET_NAMES.raid, SINGLE_SHEET_NAMES.decisions].forEach(function(name) { report.conditionalFormats[name] = ss.getSheetByName(name).getConditionalFormatRules().length > 0; });
  report.dashboardFormulas = dashboard.getRange('A5:H5').getFormulas()[0];
  report.ganttFormulas = [gantt.getRange('A5').getFormula(), gantt.getRange('L4').getFormula(), gantt.getRange('L5').getFormula()];
  report.passed = report.formulaErrors.length === 0 && report.dashboardFormulas.every(function(formula) { return formula.length > 0; }) && report.ganttFormulas.every(function(formula) { return formula.length > 0; }) && Object.keys(report.validations).every(function(key) { return report.validations[key]; }) && Object.keys(report.filters).every(function(key) { return report.filters[key]; }) && Object.keys(report.frozenHeaders).every(function(key) { return report.frozenHeaders[key]; }) && Object.keys(report.conditionalFormats).every(function(key) { return report.conditionalFormats[key]; });
  return report;
}
