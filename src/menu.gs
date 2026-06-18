/**
 * menu.gs
 * スプレッドシートを開いたときに「評価運用」メニューを追加する。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('評価運用')
    .addItem('① 評価シートを生成 / 再生成', 'menuBuild_')
    .addSeparator()
    .addItem('② 昇格・降格を判定', 'evaluatePromotionDemotion')
    .addItem('③ 当期を確定して履歴へ', 'menuCommit_')
    .addSeparator()
    .addItem('対象期を変更', 'menuSetPeriod_')
    .addToUi();
}

/** メニュー：生成（既存データがある場合は確認） */
function menuBuild_() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var has = ss.getSheetByName(SHEETS.ROSTER);
  if (has) {
    var res = ui.alert('評価シートの再生成',
      '全シートを作り直します。入力済みのデータ（名簿・粗利・評価・履歴）は消えます。続けますか？',
      ui.ButtonSet.OK_CANCEL);
    if (res !== ui.Button.OK) return;
  }
  var cfg = ss.getSheetByName(SHEETS.CONFIG);
  var period = cfg ? cfg.getRange('B44').getValue() : '';
  var resp = ui.prompt('対象期', '対象期を入力してください（例: 2026前期）', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var label = resp.getResponseText().trim() || period || DEFAULT_PERIOD;
  buildWorkbook(label);
  ui.alert('生成しました', '「名簿」シートからメンバーを入力してください。', ui.ButtonSet.OK);
}

/** メニュー：確定 */
function menuCommit_() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('当期の確定',
    '評価結果サマリの内容を履歴へ追記します。確定後に行ってください。続けますか？',
    ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  commitPeriod();
}

/** メニュー：対象期の変更 */
function menuSetPeriod_() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName(SHEETS.CONFIG);
  if (!cfg) { ui.alert('先に評価シートを生成してください。'); return; }
  var resp = ui.prompt('対象期の変更', '新しい対象期を入力（例: 2026後期）', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var label = resp.getResponseText().trim();
  if (label) cfg.getRange('B44').setValue(label);
}
