function singleBuildMaster_(sheet) {
  const rows = [
    ['状態','優先度','種別','RAID種別','確率','影響度','記録種別','変更管理状態','全体状態','開発アプローチ','フェーズ','エンゲージメント'],
    ['未着手','最優先','タスク','リスク','低','低','変更要求','起票','健全','予測型','立上げ','不認識'],
    ['進行中','高','マイルストーン','前提','中','中','意思決定','影響分析中','注意','反復型','計画','抵抗'],
    ['ブロック','中','成果物','課題','高','高','','承認','危険','漸進型','実行','中立'],
    ['完了','低','ユーザーストーリー','依存関係','','','','却下','','アジャイル','監視・コントロール','支持'],
    ['保留','','','','','','','実施中','','ハイブリッド','終結','主導'],
    ['中止','','','','','','','完了','','','','']
  ];
  sheet.clear(); sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows); styleTable_(sheet, 12, rows.length);
  sheet.setFrozenRows(1); resetFilter_(sheet, 'A1:L7');
  for (let column = 1; column <= 12; column += 1) sheet.setColumnWidth(column, 135);
}

function singleBuildGuide_(sheet) {
  const rows = [
    ['単一プロジェクト管理テンプレート', 'PMBOK第8版を参照した、現場で更新し続けられる最小構成です。'],
    ['使い方', 'Project Charterで基準線を合意し、Work Items・RAID・Stakeholders・Decision Logを定例で更新します。'],
    ['ガバナンス', 'Project Charterのスポンサー、意思決定方法、Decision Logで説明責任を残します。'],
    ['スコープ', 'Project Charterの目的・成果物・受入基準とWork Itemsで管理します。'],
    ['スケジュール', 'Work Itemsの開始日・期限・進捗、マイルストーン種別を使います。'],
    ['財務', 'Project Charterの予算ベースライン・予測・実績を定例更新します。'],
    ['ステークホルダー', 'Stakeholdersで影響度、関心、コミュニケーションを管理します。'],
    ['リソース', '担当者、チーム、Stakeholdersの役割をWork Itemsへ反映します。'],
    ['リスク', 'RAID Logでリスク、前提、課題、依存関係と対応を追跡します。'],
    ['テーラリング・価値・品質', '目的に照らして列や運用頻度を調整し、品質・価値・説明責任をDashboardとCharterで確認します。'],
    ['後続拡張', '予算・資源の明細、独立した変更台帳、定型Status ReportsはMVPでは別シートにせず、Charter・Decision Log・Work Itemsで扱います。'],
    ['PMI PMBOK 第8版', 'https://www.pmi.org/standards/pmbok'],
    ['ISO 21502', 'https://committee.iso.org/sites/tc258/home/projects/published/iso-21502.html'],
    ['Scrum Guide', 'https://scrumguides.org/scrum-guide.html']
  ];
  sheet.clear(); sheet.getRange(1, 1, rows.length, 2).setValues(rows); styleHeader_(sheet, 2); sheet.getRange('A2:B' + rows.length).setWrap(true);
  sheet.getRange('A1:B1').merge().setValue('Guide — PMBOK-informed MVP'); sheet.getRange('A1:B1').setBackground('#E8EAED').setFontWeight('bold');
  sheet.setFrozenRows(1); sheet.setColumnWidth(1, 220); sheet.setColumnWidth(2, 720);
}

function singleBuildCharter_(sheet) {
  const project = SINGLE_PROJECT_SEED.project;
  const rows = [
    ['項目', '内容'], ['Project ID', project.id], ['プロジェクト名', project.name], ['目的・期待価値', project.purpose], ['成功指標', project.successMeasures || ''], ['スポンサー', project.sponsor], ['プロジェクト責任者', project.manager], ['全体状態', project.status], ['開発アプローチ', project.deliveryApproach || ''], ['現在フェーズ', project.phase || ''], ['スコープ・主要成果物', project.scope || ''], ['スコープ外', project.outOfScope || ''], ['受入基準・品質', project.acceptance || ''], ['開始日', singleDate_(project.startDate)], ['目標完了日', singleDate_(project.targetDate)], ['予算ベースライン', project.budgetBaseline || ''], ['予算予測', project.budgetForecast || ''], ['予算実績', project.budgetActual || ''], ['意思決定・ガバナンス', project.governance || ''], ['主要制約', project.constraints || ''], ['持続可能性・ガードレール', project.sustainability || ''], ['最終レビュー日', singleDate_(project.lastReview)], ['次回レビュー日', singleDate_(project.nextReview)], ['テーラリング方針', project.tailoring || '']
  ];
  sheet.clear(); sheet.getRange(1, 1, rows.length, 2).setValues(rows); styleTable_(sheet, 2, rows.length); sheet.setFrozenRows(1);
  sheet.getRange('B14:B15').setNumberFormat('yyyy-mm-dd'); sheet.getRange('B22:B23').setNumberFormat('yyyy-mm-dd'); sheet.getRange('B16:B18').setNumberFormat('#,##0');
  sheet.getRange('B2:B24').setWrap(true); sheet.setColumnWidth(1, 210); sheet.setColumnWidth(2, 760);
  const master = singleSheet_(sheet.getParent(), 'Master');
  setListValidation_(sheet.getRange('B8'), master.getRange('I2:I4'), '全体状態を選択');
  setListValidation_(sheet.getRange('B9'), master.getRange('J2:J6'), '開発アプローチを選択');
  setListValidation_(sheet.getRange('B10'), master.getRange('K2:K6'), 'フェーズを選択');
}

function singleBuildWorkItems_(sheet) {
  const headers = ['ID','親ID','種別','タイトル','状態','優先度','担当者','開始日','期限','進捗%','見積工数(h)','実績工数(h)','マイルストーン','受入基準','タグ','URL','メモ'];
  const rows = SINGLE_PROJECT_SEED.workItems.map(function(item) { return [item.id,item.parentId || '',item.type,item.title,item.status,item.priority,item.owner,singleDate_(item.start),singleDate_(item.due),item.progress || 0,item.estimate || '',item.actual || '',item.milestone || '',item.acceptance || '',item.tags || '',item.url || '',item.memo || '']; });
  sheet.clear(); sheet.getRange(1, 1, 1, headers.length).setValues([headers]); if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows); styleTable_(sheet, headers.length, rows.length + 1);
  sheet.setFrozenRows(1); sheet.setFrozenColumns(3); resetFilter_(sheet, 'A1:Q' + SINGLE_MAX_ROWS);
  const master = singleSheet_(sheet.getParent(), 'Master'); setListValidation_(sheet.getRange('C2:C' + SINGLE_MAX_ROWS), master.getRange('C2:C4'), '種別を選択'); setListValidation_(sheet.getRange('E2:E' + SINGLE_MAX_ROWS), master.getRange('A2:A7'), '状態を選択'); setListValidation_(sheet.getRange('F2:F' + SINGLE_MAX_ROWS), master.getRange('B2:B5'), '優先度を選択');
  setDateValidation_(sheet.getRange('H2:I' + SINGLE_MAX_ROWS)); setPercentValidation_(sheet.getRange('J2:J' + SINGLE_MAX_ROWS)); sheet.getRange('H2:I' + SINGLE_MAX_ROWS).setNumberFormat('yyyy-mm-dd'); sheet.getRange('J2:J' + SINGLE_MAX_ROWS).setNumberFormat('0%'); sheet.getRange('K2:L' + SINGLE_MAX_ROWS).setNumberFormat('#,##0');
  sheet.getRange('D2:Q' + SINGLE_MAX_ROWS).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); sheet.setConditionalFormatRules([]); applyStatusRules_(sheet, 'E2:E' + SINGLE_MAX_ROWS); applyPriorityRules_(sheet, 'F2:F' + SINGLE_MAX_ROWS); applyOverdueRule_(sheet, 'I2:I' + SINGLE_MAX_ROWS, '$I2', '$E2');
  [105,95,120,280,105,85,120,105,105,80,90,90,150,240,130,260,240].forEach(function(width, index) { sheet.setColumnWidth(index + 1, width); });
}

function singleBuildRaid_(sheet) {
  const headers = ['ID','種別','名称','状態','確率','影響度','露出度','担当者','登録日','期限','トリガー・依存','対応方針','次のアクション','URL','メモ'];
  const rows = SINGLE_PROJECT_SEED.raid.map(function(item) { return [item.id,item.type,item.name,item.status,item.probability,item.impact,'',item.owner,singleDate_(item.registered),singleDate_(item.due),item.trigger || '',item.response || '',item.nextAction || '',item.url || '',item.memo || '']; });
  singleBuildLogTable_(sheet, headers, rows, 'A1:O' + SINGLE_MAX_ROWS, function(master) { setListValidation_(sheet.getRange('B2:B' + SINGLE_MAX_ROWS), master.getRange('D2:D5'), 'RAID種別を選択'); setListValidation_(sheet.getRange('D2:D' + SINGLE_MAX_ROWS), master.getRange('A2:A7'), '状態を選択'); setListValidation_(sheet.getRange('E2:F' + SINGLE_MAX_ROWS), master.getRange('E2:E4'), '確率・影響度を選択'); setDateValidation_(sheet.getRange('I2:J' + SINGLE_MAX_ROWS)); sheet.getRange('I2:J' + SINGLE_MAX_ROWS).setNumberFormat('yyyy-mm-dd'); applyStatusRules_(sheet, 'D2:D' + SINGLE_MAX_ROWS); applyOverdueRule_(sheet, 'J2:J' + SINGLE_MAX_ROWS, '$J2', '$D2'); });
  sheet.getRange('G2').setFormula('=IF($A2="","",IF(OR($D2="完了",$D2="中止"),"",IF(OR(AND($E2="高",$F2<>"低"),AND($F2="高",$E2<>"低")),"高",IF(OR($E2="中",$F2="中",$E2="高",$F2="高"),"中","低"))))');
  sheet.getRange('G2').copyTo(sheet.getRange('G2:G' + SINGLE_MAX_ROWS), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  singleApplyExposureRules_(sheet, 'G2:G' + SINGLE_MAX_ROWS);
}

function singleBuildStakeholders_(sheet) {
  const headers = ['ID','氏名・組織','役割','影響度','関心度','現状','目標','コミュニケーション','頻度','担当者','最終連絡日','次回連絡日','メモ'];
  const rows = SINGLE_PROJECT_SEED.stakeholders.map(function(item) { return [item.id,item.name,item.role,item.influence,item.interest,item.current || '',item.target || '',item.communication || '',item.cadence || '',item.owner || '',singleDate_(item.last),singleDate_(item.next),item.memo || '']; });
  singleBuildLogTable_(sheet, headers, rows, 'A1:M' + SINGLE_MAX_ROWS, function(master) { setListValidation_(sheet.getRange('D2:E' + SINGLE_MAX_ROWS), master.getRange('E2:E4'), '影響度・関心度を選択'); setListValidation_(sheet.getRange('F2:G' + SINGLE_MAX_ROWS), master.getRange('L2:L6'), 'エンゲージメントを選択'); setDateValidation_(sheet.getRange('K2:L' + SINGLE_MAX_ROWS)); sheet.getRange('K2:L' + SINGLE_MAX_ROWS).setNumberFormat('yyyy-mm-dd'); });
}

function singleBuildDecisions_(sheet) {
  const headers = ['ID','記録種別','登録日','論点','内容','理由','影響','状態','担当者・決定者','目標日','決定日','関連ID','URL','メモ'];
  const rows = SINGLE_PROJECT_SEED.changesDecisions.map(function(item) { return [item.id,item.type,singleDate_(item.registered),item.topic,item.content,item.reason || '',item.impact || '',item.status,item.owner || '',singleDate_(item.target),singleDate_(item.decisionDate),item.relatedIds || '',item.url || '',item.memo || '']; });
  singleBuildLogTable_(sheet, headers, rows, 'A1:N' + SINGLE_MAX_ROWS, function(master) { setListValidation_(sheet.getRange('B2:B' + SINGLE_MAX_ROWS), master.getRange('G2:G3'), '変更要求または意思決定を選択'); setListValidation_(sheet.getRange('H2:H' + SINGLE_MAX_ROWS), master.getRange('H2:H7'), '変更管理状態を選択'); setDateValidation_(sheet.getRange('C2:C' + SINGLE_MAX_ROWS)); setDateValidation_(sheet.getRange('J2:K' + SINGLE_MAX_ROWS)); sheet.getRange('C2:C' + SINGLE_MAX_ROWS).setNumberFormat('yyyy-mm-dd'); sheet.getRange('J2:K' + SINGLE_MAX_ROWS).setNumberFormat('yyyy-mm-dd'); singleApplyChangeRules_(sheet, 'H2:H' + SINGLE_MAX_ROWS); });
}

function singleBuildLogTable_(sheet, headers, rows, filterRange, configure) {
  sheet.clear(); sheet.getRange(1, 1, 1, headers.length).setValues([headers]); if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows); styleTable_(sheet, headers.length, rows.length + 1); sheet.setFrozenRows(1); sheet.setFrozenColumns(2); resetFilter_(sheet, filterRange); sheet.getRange(2, 1, SINGLE_MAX_ROWS - 1, headers.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); sheet.setConditionalFormatRules([]); configure(singleSheet_(sheet.getParent(), 'Master')); for (let column = 1; column <= headers.length; column += 1) sheet.setColumnWidth(column, column < 3 ? 115 : 180);
}

function singleApplyChangeRules_(sheet, a1Range) {
  const rules = sheet.getConditionalFormatRules();
  [['起票','#FEF3C7','#92400E'],['影響分析中','#E8F0FE','#174EA6'],['承認','#E6F4EA','#137333'],['却下','#FDECEC','#B3261E'],['実施中','#E8F0FE','#174EA6'],['完了','#E6F4EA','#137333']].forEach(function(spec) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(spec[0]).setBackground(spec[1]).setFontColor(spec[2]).setRanges([sheet.getRange(a1Range)]).build());
  });
  sheet.setConditionalFormatRules(rules);
}

function singleApplyExposureRules_(sheet, a1Range) {
  const rules = sheet.getConditionalFormatRules();
  [['高','#FDECEC','#B3261E'],['中','#FEF3C7','#92400E'],['低','#E6F4EA','#137333']].forEach(function(spec) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(spec[0]).setBackground(spec[1]).setFontColor(spec[2]).setRanges([sheet.getRange(a1Range)]).build());
  });
  sheet.setConditionalFormatRules(rules);
}

function singleBuildDashboard_(sheet) {
  sheet.clear(); sheet.setConditionalFormatRules([]); sheet.getRange('A1:H1').merge().setValue('単一プロジェクト Dashboard'); sheet.getRange('A2:H2').merge().setValue(SINGLE_PROJECT_SEED.project.name + ' — ' + SINGLE_PROJECT_SEED_META.source + '（' + SINGLE_PROJECT_SEED_META.generatedAt + '）');
  sheet.getRange('A4:H4').setValues([['全体状態','未完了項目','ブロック','期限超過','高露出RAID','次のMS','予算差異','次回レビュー']]);
  sheet.getRange('A5:H5').setFormulas([['=\'Project Charter\'!B8','=COUNTIFS(\'Work Items\'!A2:A1000,"<>",\'Work Items\'!E2:E1000,"<>完了",\'Work Items\'!E2:E1000,"<>中止")','=COUNTIF(\'Work Items\'!E2:E1000,"ブロック")+COUNTIF(\'RAID Log\'!D2:D1000,"ブロック")','=COUNTIFS(\'Work Items\'!I2:I1000,"<"&TODAY(),\'Work Items\'!I2:I1000,"<>",\'Work Items\'!E2:E1000,"<>完了",\'Work Items\'!E2:E1000,"<>中止")','=COUNTIFS(\'RAID Log\'!G2:G1000,"高",\'RAID Log\'!D2:D1000,"<>完了",\'RAID Log\'!D2:D1000,"<>中止")','=IFERROR(INDEX(SORT(FILTER({\'Work Items\'!D2:D1000,\'Work Items\'!I2:I1000},\'Work Items\'!C2:C1000="マイルストーン",\'Work Items\'!E2:E1000<>"完了",\'Work Items\'!I2:I1000>=TODAY()),2,TRUE),1,1),"")','=\'Project Charter\'!B17-\'Project Charter\'!B16','=\'Project Charter\'!B23']]);
  sheet.getRange('A8:D8').setValues([['次のマイルストーン','期限','状態','担当者']]); sheet.getRange('A9').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({\'Work Items\'!D2:D1000,\'Work Items\'!I2:I1000,\'Work Items\'!E2:E1000,\'Work Items\'!G2:G1000},\'Work Items\'!C2:C1000="マイルストーン",\'Work Items\'!E2:E1000<>"完了"),2,TRUE),8,4),"")'); sheet.getRange('B9:B16').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('G5').setNumberFormat('#,##0'); sheet.getRange('H5').setNumberFormat('yyyy-mm-dd');
  sheet.setHiddenGridlines(true); sheet.setFrozenRows(2); sheet.getRange('A1:H20').setFontFamily('Arial').setVerticalAlignment('middle'); sheet.getRange('A1:H1').setBackground('#E8EAED').setFontWeight('bold').setFontSize(18); sheet.getRange('A4:H4').setBackground('#E8EAED').setFontWeight('bold').setWrap(true); sheet.getRange('A5:H5').setFontWeight('bold').setFontSize(14); sheet.getRange('A8:D8').setBackground('#E8EAED').setFontWeight('bold'); [115,130,105,105,120,240,120,120].forEach(function(width, index) { sheet.setColumnWidth(index + 1, width); }); applyHealthRules_(sheet, 'A5');
}
