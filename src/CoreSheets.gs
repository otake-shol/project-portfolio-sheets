function buildMaster_(sheet) {
  const rows = [
    ['状態','優先度','担当者','カテゴリ','管理形態','種別','リスク確率','影響度','健康度','依存種別','意思決定状態','更新期限(日)'],
    ['棚卸し待ち','棚卸し待ち','棚卸し待ち','モバイル','git submodule','リスク','棚卸し待ち','棚卸し待ち','健全','前提','提案中',30],
    ['未着手','最優先','','Web','親リポジトリ直管理','マイルストーン','低','低','注意','成果物','決定',''],
    ['進行中','高','','Shopify','','','中','中','危険','環境','見直し中',''],
    ['ブロック','中','','翠進','','','高','高','','外部','撤回',''],
    ['完了','低','','キャリア','','','','','その他','','',''],
    ['保留','','','','','','','','','','',''],
    ['中止','','','','','','','','','','','']
  ];
  sheet.clear();
  sheet.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  styleTable_(sheet, 12, rows.length);
  sheet.getRange('A1').setNote('Project Plan Workspaceの状態モデルを再利用し、横断棚卸し用の候補を追加。');
  sheet.getRange('F1').setNote('Work Itemのmilestone/impedimentの考え方を横断管理向けに集約。');
  sheet.getRange('L1').setNote('活動中プロジェクトを更新漏れと判定する日数。初期値は30日。');
  for (let column = 1; column <= 12; column++) sheet.setColumnWidth(column, 135);
  sheet.setFrozenRows(1);
  resetFilter_(sheet, 'A1:L8');
}

function upgradeMaster_(sheet) {
  if (sheet.getRange('A1').getValue() === '') buildMaster_(sheet);
  sheet.getRange('I1:L1').setValues([['健康度','依存種別','意思決定状態','更新期限(日)']]);
  sheet.getRange('I2:I4').setValues([['健全'],['注意'],['危険']]);
  sheet.getRange('J2:J6').setValues([['前提'],['成果物'],['環境'],['外部'],['その他']]);
  sheet.getRange('K2:K5').setValues([['提案中'],['決定'],['見直し中'],['撤回']]);
  if (sheet.getRange('L2').getValue() === '') sheet.getRange('L2').setValue(30);
  styleHeader_(sheet, 12);
  sheet.getRange('L1').setNote('活動中プロジェクトを更新漏れと判定する日数。初期値は30日。');
  for (let column = 9; column <= 12; column++) sheet.setColumnWidth(column, 135);
  sheet.setFrozenRows(1);
  resetFilter_(sheet, 'A1:L8');
}

function buildProjects_(sheet, rows, master) {
  const headers = ['Project ID','カテゴリ','名称','ローカルパス','管理形態','状態','優先度','担当者','次のアクション','期限','Repository URL','メモ','最終確認日','健康度','健康度理由'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.getRange(2,1,rows.length,13).setValues(rows);
  styleTable_(sheet, headers.length, rows.length + 1);
  configureProjects_(sheet, master);
}

function upgradeProjects_(sheet, master) {
  if (sheet.getRange('A1').getValue() === '') {
    throw new Error('ProjectsシートにProject ID列がありません');
  }
  sheet.getRange('N1:O1').setValues([['健康度','健康度理由']]);
  styleHeader_(sheet, 15);
  configureProjects_(sheet, master);
}

function configureProjects_(sheet, master) {
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  resetFilter_(sheet, 'A1:O' + MAX_ROWS);
  setListValidation_(sheet.getRange('B2:B' + MAX_ROWS), master.getRange('D2:D6'), 'Masterのカテゴリ候補から選択');
  setListValidation_(sheet.getRange('E2:E' + MAX_ROWS), master.getRange('E2:E3'), 'Masterの管理形態から選択');
  setListValidation_(sheet.getRange('F2:F' + MAX_ROWS), master.getRange('A2:A8'), 'Masterの状態から選択');
  setListValidation_(sheet.getRange('G2:G' + MAX_ROWS), master.getRange('B2:B6'), 'Masterの優先度から選択');
  setListValidation_(sheet.getRange('H2:H' + MAX_ROWS), master.getRange('C2:C' + MAX_ROWS), '担当者候補をMasterへ追加して選択');
  setDateValidation_(sheet.getRange('J2:J' + MAX_ROWS));
  setDateValidation_(sheet.getRange('M2:M' + MAX_ROWS));
  sheet.getRange('J2:J' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('M2:M' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('N2').setFormula(projectHealthFormula_(2));
  sheet.getRange('N2').copyTo(sheet.getRange('N2:N' + MAX_ROWS), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  sheet.getRange('O2').setFormula(projectHealthReasonFormula_(2));
  sheet.getRange('O2').copyTo(sheet.getRange('O2:O' + MAX_ROWS), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  sheet.getRange('D2:D' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('I2:I' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('K2:L' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('O2:O' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setConditionalFormatRules([]);
  applyStatusRules_(sheet, 'F2:F' + MAX_ROWS);
  applyPriorityRules_(sheet, 'G2:G' + MAX_ROWS);
  applyOverdueRule_(sheet, 'J2:J' + MAX_ROWS, '$J2', '$F2');
  applyHealthRules_(sheet, 'N2:N' + MAX_ROWS);
  applyStaleRule_(sheet, 'M2:M' + MAX_ROWS);
  [95,100,180,420,120,110,90,120,240,105,300,260,105,90,260].forEach(function(width,index) {
    sheet.setColumnWidth(index + 1, width);
  });
}

function projectHealthFormula_(row) {
  return '=IF($A' + row + '="","",IF(OR($F' + row + '="ブロック",COUNTIFS(Dependencies!$B$2:$B$1000,$A' + row + ',Dependencies!$E$2:$E$1000,"ブロック")>0,COUNTIFS(\'Risks & Milestones\'!$B$2:$B$1000,$A' + row + ',\'Risks & Milestones\'!$A$2:$A$1000,"リスク",\'Risks & Milestones\'!$D$2:$D$1000,"ブロック")>0),"危険",IF(OR($F' + row + '="棚卸し待ち",AND($J' + row + '<>"",$J' + row + '<TODAY(),$F' + row + '<>"完了",$F' + row + '<>"中止"),AND($F' + row + '<>"完了",$F' + row + '<>"中止",OR($M' + row + '="",TODAY()-$M' + row + '>Master!$L$2))),"注意","健全")))';
}

function projectHealthReasonFormula_(row) {
  return '=IF($A' + row + '="","",TEXTJOIN(" / ",TRUE,IF($F' + row + '="ブロック","状態がブロック",""),IF(COUNTIFS(Dependencies!$B$2:$B$1000,$A' + row + ',Dependencies!$E$2:$E$1000,"ブロック")>0,"依存関係にブロック",""),IF(COUNTIFS(\'Risks & Milestones\'!$B$2:$B$1000,$A' + row + ',\'Risks & Milestones\'!$A$2:$A$1000,"リスク",\'Risks & Milestones\'!$D$2:$D$1000,"ブロック")>0,"リスクがブロック",""),IF(AND($J' + row + '<>"",$J' + row + '<TODAY(),$F' + row + '<>"完了",$F' + row + '<>"中止"),"期限超過",""),IF(AND($F' + row + '<>"完了",$F' + row + '<>"中止",$M' + row + '=""),"最終確認日なし",""),IF(AND($F' + row + '<>"完了",$F' + row + '<>"中止",$M' + row + '<>"",TODAY()-$M' + row + '>Master!$L$2),"更新期限超過",""),IF($F' + row + '="棚卸し待ち","棚卸し待ち","")))';
}

function buildTasks_(sheet, projects, master, preserveData) {
  const headers = ['Task ID','Project ID','タイトル','状態','優先度','担当者','開始日','期限','Milestone','タグ','外部URL','メモ'];
  if (!preserveData) sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  resetFilter_(sheet, 'A1:L' + MAX_ROWS);
  setListValidation_(sheet.getRange('B2:B' + MAX_ROWS), projects.getRange('A2:A' + MAX_ROWS), 'ProjectsのProject IDから選択');
  setListValidation_(sheet.getRange('D2:D' + MAX_ROWS), master.getRange('A2:A8'), 'Masterの状態から選択');
  setListValidation_(sheet.getRange('E2:E' + MAX_ROWS), master.getRange('B2:B6'), 'Masterの優先度から選択');
  setListValidation_(sheet.getRange('F2:F' + MAX_ROWS), master.getRange('C2:C' + MAX_ROWS), '担当者候補をMasterへ追加して選択');
  setDateValidation_(sheet.getRange('G2:H' + MAX_ROWS));
  sheet.getRange('G2:H' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('C2:C' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('J2:L' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setConditionalFormatRules([]);
  applyStatusRules_(sheet, 'D2:D' + MAX_ROWS);
  applyPriorityRules_(sheet, 'E2:E' + MAX_ROWS);
  applyOverdueRule_(sheet, 'H2:H' + MAX_ROWS, '$H2', '$D2');
  [95,95,280,110,90,120,105,105,160,180,280,240].forEach(function(width,index) {
    sheet.setColumnWidth(index + 1, width);
  });
}

function buildRisks_(sheet, projects, master, preserveData) {
  const headers = ['種別','Project ID','名称','状態','確率','影響度','担当者','期限','対応方針','開始日','進捗%','外部URL'];
  if (!preserveData) sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  resetFilter_(sheet, 'A1:L' + MAX_ROWS);
  setListValidation_(sheet.getRange('A2:A' + MAX_ROWS), master.getRange('F2:F3'), 'リスクまたはマイルストーンを選択');
  setListValidation_(sheet.getRange('B2:B' + MAX_ROWS), projects.getRange('A2:A' + MAX_ROWS), 'ProjectsのProject IDから選択');
  setListValidation_(sheet.getRange('D2:D' + MAX_ROWS), master.getRange('A2:A8'), 'Masterの状態から選択');
  setListValidation_(sheet.getRange('E2:E' + MAX_ROWS), master.getRange('G2:G5'), 'Masterのリスク確率から選択');
  setListValidation_(sheet.getRange('F2:F' + MAX_ROWS), master.getRange('H2:H5'), 'Masterの影響度から選択');
  setListValidation_(sheet.getRange('G2:G' + MAX_ROWS), master.getRange('C2:C' + MAX_ROWS), '担当者候補をMasterへ追加して選択');
  setDateValidation_(sheet.getRange('H2:H' + MAX_ROWS));
  setDateValidation_(sheet.getRange('J2:J' + MAX_ROWS));
  setPercentValidation_(sheet.getRange('K2:K' + MAX_ROWS));
  sheet.getRange('H2:H' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('J2:J' + MAX_ROWS).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('K2:K' + MAX_ROWS).setNumberFormat('0%');
  sheet.getRange('C2:C' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange('I2:I' + MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setConditionalFormatRules([]);
  applyStatusRules_(sheet, 'D2:D' + MAX_ROWS);
  applyOverdueRule_(sheet, 'H2:H' + MAX_ROWS, '$H2', '$D2');
  [120,95,260,110,90,90,120,105,320,105,90,280].forEach(function(width,index) {
    sheet.setColumnWidth(index + 1, width);
  });
}
