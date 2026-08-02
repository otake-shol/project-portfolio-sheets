function styleTable_(sheet, columnCount, rowCount) {
  sheet.setHiddenGridlines(true);
  sheet.getRange(1,1,Math.max(rowCount,2),columnCount).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  styleHeader_(sheet, columnCount);
  sheet.setRowHeight(1, 36);
  sheet.getRange(2,1,Math.max(rowCount - 1,1),columnCount).setBackground('#FFFFFF');
}

function styleHeader_(sheet, columnCount) {
  sheet.setHiddenGridlines(true);
  sheet.getRange(1,1,1,columnCount)
    .setBackground('#E8EAED')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor('#202124')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(1, 36);
}

function resetFilter_(sheet, a1Range) {
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(a1Range).createFilter();
}

function setListValidation_(range, sourceRange, helpText) {
  range.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .setHelpText(helpText)
    .build());
}

function setDateValidation_(range) {
  range.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('日付を入力（yyyy-mm-dd）')
    .build());
}

function setPercentValidation_(range) {
  range.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 1)
    .setAllowInvalid(false)
    .setHelpText('0%から100%の範囲で入力')
    .build());
}

function applyStatusRules_(sheet, a1Range) {
  const range = sheet.getRange(a1Range);
  const rules = sheet.getConditionalFormatRules();
  [
    ['棚卸し待ち','#FEF3C7','#92400E'],
    ['未着手','#F1F3F4','#3C4043'],
    ['進行中','#E8F0FE','#174EA6'],
    ['ブロック','#FDECEC','#B3261E'],
    ['完了','#E6F4EA','#137333'],
    ['保留','#F3E8FD','#681DA8'],
    ['中止','#F1F3F4','#5F6368']
  ].forEach(function(spec) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(spec[0])
      .setBackground(spec[1])
      .setFontColor(spec[2])
      .setRanges([range])
      .build());
  });
  sheet.setConditionalFormatRules(rules);
}

function applyPriorityRules_(sheet, a1Range) {
  const range = sheet.getRange(a1Range);
  const rules = sheet.getConditionalFormatRules();
  [
    ['棚卸し待ち','#FEF3C7','#92400E'],
    ['最優先','#FDECEC','#B3261E'],
    ['高','#FCE8E6','#B06000'],
    ['中','#FFF7E0','#7A5901'],
    ['低','#F1F3F4','#5F6368']
  ].forEach(function(spec) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(spec[0])
      .setBackground(spec[1])
      .setFontColor(spec[2])
      .setRanges([range])
      .build());
  });
  sheet.setConditionalFormatRules(rules);
}

function applyHealthRules_(sheet, a1Range) {
  const range = sheet.getRange(a1Range);
  const rules = sheet.getConditionalFormatRules();
  [
    ['健全','#E6F4EA','#137333'],
    ['注意','#FEF3C7','#92400E'],
    ['危険','#FDECEC','#B3261E']
  ].forEach(function(spec) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(spec[0])
      .setBackground(spec[1])
      .setFontColor(spec[2])
      .setRanges([range])
      .build());
  });
  sheet.setConditionalFormatRules(rules);
}

function applyDecisionRules_(sheet, a1Range) {
  const range = sheet.getRange(a1Range);
  const rules = sheet.getConditionalFormatRules();
  [
    ['提案中','#FEF3C7','#92400E'],
    ['決定','#E6F4EA','#137333'],
    ['見直し中','#E8F0FE','#174EA6'],
    ['撤回','#F1F3F4','#5F6368']
  ].forEach(function(spec) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(spec[0])
      .setBackground(spec[1])
      .setFontColor(spec[2])
      .setRanges([range])
      .build());
  });
  sheet.setConditionalFormatRules(rules);
}

function applyOverdueRule_(sheet, a1Range, dateRef, stateRef) {
  const range = sheet.getRange(a1Range);
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + dateRef + '<TODAY(),' + dateRef + '<>"",' + stateRef + '<>"完了",' + stateRef + '<>"中止")')
    .setBackground('#FDECEC')
    .setFontColor('#B3261E')
    .setRanges([range])
    .build());
  sheet.setConditionalFormatRules(rules);
}

function applyStaleRule_(sheet, a1Range) {
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",REGEXMATCH($O2,"最終確認日なし|更新期限超過"))')
    .setBackground('#FEF3C7')
    .setFontColor('#92400E')
    .setRanges([sheet.getRange(a1Range)])
    .build());
  sheet.setConditionalFormatRules(rules);
}

function applySelfDependencyRule_(sheet, a1Range) {
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"",$B2=$C2)')
    .setBackground('#FDECEC')
    .setFontColor('#B3261E')
    .setRanges([sheet.getRange(a1Range)])
    .build());
  sheet.setConditionalFormatRules(rules);
}

function applyReviewDueRule_(sheet, a1Range) {
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$J2<>"",$J2<TODAY(),$I2<>"撤回")')
    .setBackground('#FEF3C7')
    .setFontColor('#92400E')
    .setRanges([sheet.getRange(a1Range)])
    .build());
  sheet.setConditionalFormatRules(rules);
}

function applyTimelineDueRule_(sheet, a1Range) {
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$F2<>"",$F2<TODAY(),$C2<>"完了",$C2<>"中止")')
    .setBackground('#FDECEC')
    .setFontColor('#B3261E')
    .setRanges([sheet.getRange(a1Range)])
    .build());
  sheet.setConditionalFormatRules(rules);
}
