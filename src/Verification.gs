function verifyWorkbook_(ss) {
  SpreadsheetApp.flush();
  const expectedSheets = [
    'Dashboard', 'Projects', 'Tasks', 'Risks & Milestones', 'Dependencies',
    'Milestone Timeline', 'Decision Log', 'Master'
  ];
  const formulaErrors = [];
  ss.getSheets().forEach(function(sheet) {
    sheet.getDataRange().getDisplayValues().forEach(function(row, rowIndex) {
      row.forEach(function(value, columnIndex) {
        if (/^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|ERROR!)/.test(value)) {
          formulaErrors.push(sheet.getName() + '!' + sheet.getRange(rowIndex + 1,columnIndex + 1).getA1Notation() + ':' + value);
        }
      });
    });
  });

  const projects = ss.getSheetByName('Projects');
  const tasks = ss.getSheetByName('Tasks');
  const risks = ss.getSheetByName('Risks & Milestones');
  const dependencies = ss.getSheetByName('Dependencies');
  const timeline = ss.getSheetByName('Milestone Timeline');
  const decisions = ss.getSheetByName('Decision Log');
  const dashboard = ss.getSheetByName('Dashboard');
  const master = ss.getSheetByName('Master');
  const projectIds = projects.getRange('A2:A' + MAX_ROWS).getValues()
    .map(function(row) { return row[0]; })
    .filter(function(value) { return value !== ''; });
  const projectIdSet = new Set(projectIds);
  const invalidProjectRefs = [];

  [
    [tasks, 'B'],
    [risks, 'B'],
    [dependencies, 'B'],
    [dependencies, 'C'],
    [decisions, 'B']
  ].forEach(function(spec) {
    const sheet = spec[0];
    const column = spec[1];
    sheet.getRange(column + '2:' + column + MAX_ROWS).getValues().forEach(function(row, index) {
      const value = row[0];
      if (value !== '' && !projectIdSet.has(value)) invalidProjectRefs.push(sheet.getName() + '!' + column + (index + 2) + ':' + value);
    });
  });

  const selfDependencies = dependencies.getRange('B2:C' + MAX_ROWS).getValues()
    .map(function(row,index) { return row[0] !== '' && row[0] === row[1] ? 'Dependencies!B' + (index + 2) + ':C' + (index + 2) : ''; })
    .filter(function(value) { return value !== ''; });

  const validationChecks = {
    projectsStatus: Boolean(projects.getRange('F2').getDataValidation()),
    projectsPriority: Boolean(projects.getRange('G2').getDataValidation()),
    tasksProjectId: Boolean(tasks.getRange('B2').getDataValidation()),
    risksType: Boolean(risks.getRange('A2').getDataValidation()),
    milestoneProgress: Boolean(risks.getRange('K2').getDataValidation()),
    dependencyProjectId: Boolean(dependencies.getRange('B2').getDataValidation()),
    dependencyType: Boolean(dependencies.getRange('D2').getDataValidation()),
    decisionProjectId: Boolean(decisions.getRange('B2').getDataValidation()),
    decisionStatus: Boolean(decisions.getRange('I2').getDataValidation())
  };
  const sheetNames = ss.getSheets().map(function(sheet) { return sheet.getName(); });
  const report = {
    title: ss.getName(),
    sheets: sheetNames,
    missingSheets: expectedSheets.filter(function(name) { return sheetNames.indexOf(name) === -1; }),
    projectRows: projectIds.length,
    dashboardLabels: dashboard.getRange('A4:H4').getDisplayValues()[0],
    dashboardValues: dashboard.getRange('A5:H5').getDisplayValues()[0],
    dashboardFormulas: dashboard.getRange('A5:H5').getFormulas()[0],
    projectHealthFormulas: projects.getRange('N2:O2').getFormulas()[0],
    milestoneTimelineFormula: timeline.getRange('A2').getFormula(),
    staleDays: master.getRange('L2').getValue(),
    formulaErrors: formulaErrors,
    invalidProjectRefs: invalidProjectRefs,
    selfDependencies: selfDependencies,
    validationChecks: validationChecks,
    filters: {
      projects: Boolean(projects.getFilter()),
      tasks: Boolean(tasks.getFilter()),
      risks: Boolean(risks.getFilter()),
      dependencies: Boolean(dependencies.getFilter()),
      timeline: Boolean(timeline.getFilter()),
      decisions: Boolean(decisions.getFilter()),
      master: Boolean(master.getFilter())
    },
    frozenRows: ss.getSheets().reduce(function(result,sheet) {
      result[sheet.getName()] = sheet.getFrozenRows();
      return result;
    }, {}),
    conditionalRuleCounts: ss.getSheets().reduce(function(result,sheet) {
      result[sheet.getName()] = sheet.getConditionalFormatRules().length;
      return result;
    }, {})
  };
  report.passed = report.missingSheets.length === 0
    && report.formulaErrors.length === 0
    && report.invalidProjectRefs.length === 0
    && report.selfDependencies.length === 0
    && Object.keys(validationChecks).every(function(key) { return validationChecks[key]; });
  return report;
}
