function buildDependencies_(sheet, projects, master, preserveData) {
  const headers = ['Dependency ID','Project ID','依存先Project ID','種別','状態','影響度','担当者','解消予定日','次のアクション','メモ'];
  if (!preserveData) sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  resetFilter_(sheet, 'A1:J' + MAX_ROWS);
  setListValidation_(sheet.getRange('B2:C' + MAX_ROWS), projects.getRange('A2:A' + MAX_ROWS), 'ProjectsのProject IDから選択');
  setListValidation_(sheet.getRange('D2:D' + MAX_ROWS), master.getRange('J2:J6'), 'Masterの依存種別から選択');
  setListValidation_(sheet.getRange('E2:E' + MAX_ROWS), master.getRange('A2:A8'), 'Masterの状態から選択');
  setListValidation_(sheet.getRange('F2:F' + MAX_ROWS), master.getRange('H2:H5'), 'Masterの影響度から選択');
  setListValidation_(sheet.getRange('G2:G' + MAX_ROWS), master.getRange('C2:C' + MAX_ROWS), '担当者候補をMasterへ追加して選択');
  setDateValidation_(sheet.getRange('H2:H' + MAX_ROWS));
  sheet.getRange('H2:H' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('I2:J' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setConditionalFormatRules([]);
  applyStatusRules_(sheet, 'E2:E' + MAX_ROWS);
  applyOverdueRule_(sheet, 'H2:H' + MAX_ROWS, '$H2', '$E2');
  applySelfDependencyRule_(sheet, 'B2:C' + MAX_ROWS);
  [110,95,130,110,110,90,120,105,260,240].forEach(function(width,index) {
    sheet.setColumnWidth(index + 1, width);
  });
}

function buildDecisionLog_(sheet, projects, master, preserveData) {
  const headers = ['Decision ID','Project ID','決定日','論点','決定内容','根拠','決定者','影響範囲','状態','見直し日','関連URL','メモ'];
  if (!preserveData) sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  resetFilter_(sheet, 'A1:L' + MAX_ROWS);
  setListValidation_(sheet.getRange('B2:B' + MAX_ROWS), projects.getRange('A2:A' + MAX_ROWS), 'ProjectsのProject IDから選択');
  setDateValidation_(sheet.getRange('C2:C' + MAX_ROWS));
  setListValidation_(sheet.getRange('G2:G' + MAX_ROWS), master.getRange('C2:C' + MAX_ROWS), '担当者候補をMasterへ追加して選択');
  setListValidation_(sheet.getRange('I2:I' + MAX_ROWS), master.getRange('K2:K5'), 'Masterの意思決定状態から選択');
  setDateValidation_(sheet.getRange('J2:J' + MAX_ROWS));
  sheet.getRange('C2:C' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('J2:J' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('D2:F' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('H2:L' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setConditionalFormatRules([]);
  applyDecisionRules_(sheet, 'I2:I' + MAX_ROWS);
  applyReviewDueRule_(sheet, 'J2:J' + MAX_ROWS);
  [105,95,105,220,300,260,120,220,110,105,280,240].forEach(function(width,index) {
    sheet.setColumnWidth(index + 1, width);
  });
}

function buildMilestoneTimeline_(sheet) {
  const headers = ['Project ID','名称','状態','担当者','開始日','期限','進捗%','残日数','外部URL','対応方針'];
  sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  sheet.getRange('A2').setFormula('=IFERROR(SORT(FILTER({\'Risks & Milestones\'!B2:B1000,\'Risks & Milestones\'!C2:C1000,\'Risks & Milestones\'!D2:D1000,\'Risks & Milestones\'!G2:G1000,\'Risks & Milestones\'!J2:J1000,\'Risks & Milestones\'!H2:H1000,\'Risks & Milestones\'!K2:K1000,\'Risks & Milestones\'!H2:H1000-TODAY(),\'Risks & Milestones\'!L2:L1000,\'Risks & Milestones\'!I2:I1000},\'Risks & Milestones\'!A2:A1000="マイルストーン"),6,TRUE),"")');
  sheet.getRange('E2:F' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('G2:G' + MAX_ROWS).setNumberFormat('0%');
  sheet.getRange('H2:H' + MAX_ROWS).setNumberFormat('0');
  sheet.getRange('B2:B' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('I2:J' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  resetFilter_(sheet, 'A1:J' + MAX_ROWS);
  sheet.setConditionalFormatRules([]);
  applyStatusRules_(sheet, 'C2:C' + MAX_ROWS);
  applyTimelineDueRule_(sheet, 'F2:H' + MAX_ROWS);
  [95,260,110,120,105,105,90,90,280,320].forEach(function(width,index) {
    sheet.setColumnWidth(index + 1, width);
  });
}

function buildDashboard_(sheet, seedMeta) {
  sheet.getRange('A1:J40').breakApart().clear();
  sheet.setConditionalFormatRules([]);
  sheet.getRange('A1:I1').merge().setValue('プロジェクト横断 Dashboard');
  sheet.getRange('A2:I2').merge().setValue('Projects / Tasks / Risks / Dependencies / Decisions を数式で集計');
  sheet.getRange('A4:I4').setValues([['全プロジェクト数','進行中','要確認','期限超過','ブロック中','危険','更新漏れ','7日以内MS','30日以内MS']]);
  sheet.getRange('A5:I5').setFormulas([[
    '=COUNTA(Projects!$A$2:$A$1000)',
    '=COUNTIF(Projects!$F$2:$F$1000,"進行中")+COUNTIF(Tasks!$D$2:$D$1000,"進行中")+COUNTIF(\'Risks & Milestones\'!$D$2:$D$1000,"進行中")+COUNTIF(Dependencies!$E$2:$E$1000,"進行中")',
    '=COUNTIF(Projects!$F$2:$F$1000,"棚卸し待ち")+COUNTIF(Tasks!$D$2:$D$1000,"棚卸し待ち")+COUNTIF(\'Risks & Milestones\'!$D$2:$D$1000,"棚卸し待ち")+COUNTIF(Dependencies!$E$2:$E$1000,"棚卸し待ち")',
    '=COUNTIFS(Projects!$J$2:$J$1000,"<"&TODAY(),Projects!$J$2:$J$1000,"<>",Projects!$F$2:$F$1000,"<>完了",Projects!$F$2:$F$1000,"<>中止")+COUNTIFS(Tasks!$H$2:$H$1000,"<"&TODAY(),Tasks!$H$2:$H$1000,"<>",Tasks!$D$2:$D$1000,"<>完了",Tasks!$D$2:$D$1000,"<>中止")+COUNTIFS(\'Risks & Milestones\'!$H$2:$H$1000,"<"&TODAY(),\'Risks & Milestones\'!$H$2:$H$1000,"<>",\'Risks & Milestones\'!$D$2:$D$1000,"<>完了",\'Risks & Milestones\'!$D$2:$D$1000,"<>中止")+COUNTIFS(Dependencies!$H$2:$H$1000,"<"&TODAY(),Dependencies!$H$2:$H$1000,"<>",Dependencies!$E$2:$E$1000,"<>完了",Dependencies!$E$2:$E$1000,"<>中止")',
    '=COUNTIF(Projects!$F$2:$F$1000,"ブロック")+COUNTIF(Tasks!$D$2:$D$1000,"ブロック")+COUNTIF(\'Risks & Milestones\'!$D$2:$D$1000,"ブロック")+COUNTIF(Dependencies!$E$2:$E$1000,"ブロック")',
    '=COUNTIF(Projects!$N$2:$N$1000,"危険")',
    '=COUNTIFS(Projects!$A$2:$A$1000,"<>",Projects!$F$2:$F$1000,"<>完了",Projects!$F$2:$F$1000,"<>中止",Projects!$M$2:$M$1000,"<"&TODAY()-Master!$L$2)+COUNTIFS(Projects!$A$2:$A$1000,"<>",Projects!$F$2:$F$1000,"<>完了",Projects!$F$2:$F$1000,"<>中止",Projects!$M$2:$M$1000,"")',
    '=COUNTIFS(\'Risks & Milestones\'!$A$2:$A$1000,"マイルストーン",\'Risks & Milestones\'!$H$2:$H$1000,">="&TODAY(),\'Risks & Milestones\'!$H$2:$H$1000,"<="&TODAY()+7,\'Risks & Milestones\'!$D$2:$D$1000,"<>完了",\'Risks & Milestones\'!$D$2:$D$1000,"<>中止")',
    '=COUNTIFS(\'Risks & Milestones\'!$A$2:$A$1000,"マイルストーン",\'Risks & Milestones\'!$H$2:$H$1000,">="&TODAY(),\'Risks & Milestones\'!$H$2:$H$1000,"<="&TODAY()+30,\'Risks & Milestones\'!$D$2:$D$1000,"<>完了",\'Risks & Milestones\'!$D$2:$D$1000,"<>中止")'
  ]]);

  sheet.getRange('A8:B8').setValues([['カテゴリ','プロジェクト数']]);
  sheet.getRange('A9:A13').setFormulas([["=Master!D2"],["=Master!D3"],["=Master!D4"],["=Master!D5"],["=Master!D6"]]);
  for (let row = 9; row <= 13; row++) sheet.getRange(row,2).setFormula('=COUNTIF(Projects!$B$2:$B$1000,$A' + row + ')');

  sheet.getRange('D8:E8').setValues([['健康度','プロジェクト数']]);
  sheet.getRange('D9:D11').setFormulas([["=Master!I2"],["=Master!I3"],["=Master!I4"]]);
  for (let row = 9; row <= 11; row++) sheet.getRange(row,5).setFormula('=COUNTIF(Projects!$N$2:$N$1000,$D' + row + ')');

  sheet.getRange('G8:J8').setValues([['Project ID','直近マイルストーン','状態','期限']]);
  sheet.getRange('G9').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({\'Risks & Milestones\'!B2:B1000,\'Risks & Milestones\'!C2:C1000,\'Risks & Milestones\'!D2:D1000,\'Risks & Milestones\'!H2:H1000},\'Risks & Milestones\'!A2:A1000="マイルストーン",\'Risks & Milestones\'!H2:H1000>=TODAY(),\'Risks & Milestones\'!D2:D1000<>"完了",\'Risks & Milestones\'!D2:D1000<>"中止"),4,TRUE),7,4),"")');
  sheet.getRange('J9:J15').setNumberFormat('yyyy-mm-dd');

  sheet.getRange('A17:D17').setValues([['Project ID','更新要確認プロジェクト','最終確認日','理由']]);
  sheet.getRange('A18').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({Projects!A2:A1000,Projects!C2:C1000,Projects!M2:M1000,Projects!O2:O1000},Projects!A2:A1000<>"",Projects!F2:F1000<>"完了",Projects!F2:F1000<>"中止",((Projects!M2:M1000="")+(Projects!M2:M1000<TODAY()-Master!L2))>0),3,TRUE),10,4),"")');
  sheet.getRange('C18:C27').setNumberFormat('yyyy-mm-dd');

  sheet.getRange('G17:J17').setValues([['決定日','Project ID','最近の意思決定','状態']]);
  sheet.getRange('G18').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({\'Decision Log\'!C2:C1000,\'Decision Log\'!B2:B1000,\'Decision Log\'!D2:D1000,\'Decision Log\'!I2:I1000},\'Decision Log\'!B2:B1000<>""),1,FALSE),10,4),"")');
  sheet.getRange('G18:G27').setNumberFormat('yyyy-mm-dd');

  sheet.getRange('A31:I31').merge().setValue('初期投入: ' + seedMeta.source + ' の' + seedMeta.projectCount + '件（' + seedMeta.generatedAt + '確認）');
  sheet.getRange('A32:I32').merge().setValue('健康度と更新漏れは数式判定。更新期限はMaster!L2で変更できます。');

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(2);
  sheet.getRange('A1:J32').setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.getRange('A1:I1').setFontSize(18).setFontWeight('bold').setBackground('#F1F3F4').setHorizontalAlignment('left');
  sheet.getRange('A2:I2').setFontSize(10).setFontColor('#5F6368');
  sheet.getRange('A4:I4').setBackground('#E8EAED').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sheet.getRange('A5:I5').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('0');
  ['A8:B8','D8:E8','G8:J8','A17:D17','G17:J17'].forEach(function(a1) {
    sheet.getRange(a1).setBackground('#E8EAED').setFontWeight('bold').setWrap(true);
  });
  sheet.getRange('A31:I32').setFontSize(9).setFontColor('#5F6368').setWrap(true);
  [120,170,120,180,120,35,105,220,110,110].forEach(function(width,index) {
    sheet.setColumnWidth(index + 1, width);
  });
  sheet.setRowHeight(1, 38);
  sheet.setRowHeight(5, 34);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#FDECEC').setFontColor('#B3261E').setRanges([sheet.getRange('D5:F5')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#FEF3C7').setFontColor('#92400E').setRanges([sheet.getRange('C5'),sheet.getRange('G5:I5')]).build()
  ]);
}
