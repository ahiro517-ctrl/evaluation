/**
 * build.gs
 * 評価シート（スプレッドシート）の生成。アクティブなスプレッドシートに
 * 全シートを作成し、見出し・マスタ・計算式・入力規則を流し込む。
 *
 * メニュー「評価運用 → 評価シートを生成 / 再生成」から呼ぶ。
 */

/** 列番号→A1の列文字（1→A, 27→AA, 29→AC …） */
function colLetter_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** シートを取得（無ければ作成）し、内容・入力規則をクリアして返す */
function resetSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.clearNotes();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  return sh;
}

/**
 * 列ごとの数式関数の配列から、データ行に数式グリッドを一括投入する。
 * colFns[c] は (rowNum)=>"数式文字列" の関数。null の列は空（入力用）。
 */
function applyGrid_(sheet, startRow, numRows, colFns) {
  var ncols = colFns.length;
  var grid = [];
  for (var i = 0; i < numRows; i++) {
    var r = startRow + i;
    var line = [];
    for (var c = 0; c < ncols; c++) {
      line.push(colFns[c] ? colFns[c](r) : '');
    }
    grid.push(line);
  }
  sheet.getRange(startRow, 1, numRows, ncols).setFormulas(grid);
}

/** ヘッダー行を書いて体裁を整える */
function writeHeader_(sheet, headers) {
  var rng = sheet.getRange(1, 1, 1, headers.length);
  rng.setValues([headers]);
  rng.setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff')
     .setVerticalAlignment('middle').setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

/** メインエントリ：アクティブなスプレッドシートに評価シートを生成 */
function buildWorkbook(periodLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('アクティブなスプレッドシートがありません。');
  periodLabel = periodLabel || DEFAULT_PERIOD;

  buildConfigSheet_(ss, periodLabel);
  buildRosterSheet_(ss);
  buildGpSheet_(ss);
  buildValueSheet_(ss);
  buildSkillSheet_(ss);
  buildSummarySheet_(ss);
  buildPromoteSheet_(ss);
  buildHistorySheet_(ss);

  // 既定の空シートを掃除
  var first = ss.getSheets()[0];
  if (first.getName() === 'シート1' || first.getName() === 'Sheet1') {
    if (ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  // 設定シートを先頭へ
  ss.setActiveSheet(ss.getSheetByName(SHEETS.ROSTER));
  ss.moveActiveSheet(2);
  ss.setActiveSheet(ss.getSheetByName(SHEETS.CONFIG));
  ss.moveActiveSheet(1);

  return ss.getUrl();
}

/* ============================ 設定シート ============================ */
function buildConfigSheet_(ss, periodLabel) {
  var sh = resetSheet_(ss, SHEETS.CONFIG);

  // ── 等級基準テーブル（A1〜H11）
  sh.getRange('A1').setValue('■ 等級基準テーブル').setFontWeight('bold');
  var gHead = ['等級キー', '等級', '区分', '月収下限', '月収上限', '半期基準粗利(1.0倍)', '半期目標粗利(3.0倍)', '等級グループ'];
  sh.getRange(2, 1, 1, gHead.length).setValues([gHead]).setFontWeight('bold').setBackground('#cfd8dc');
  sh.getRange(3, 1, GRADES.length, 8).setValues(GRADES);
  sh.getRange(3, 4, GRADES.length, 4).setNumberFormat('#,##0');

  // ── 粗利判定（A13〜C21）
  sh.getRange('A13').setValue('■ 粗利判定（倍率→判定）※LOOKUP用に倍率下限の昇順').setFontWeight('bold');
  sh.getRange(14, 1, 1, 3).setValues([['倍率下限', '判定', '状態']]).setFontWeight('bold').setBackground('#cfd8dc');
  sh.getRange(15, 1, GP_JUDGE.length, 3).setValues(GP_JUDGE);
  sh.getRange(15, 1, GP_JUDGE.length, 1).setNumberFormat('0.0');

  // ── バリュー判定（A23〜C27）
  sh.getRange('A23').setValue('■ バリュー判定（合計27点満点）').setFontWeight('bold');
  sh.getRange(24, 1, 1, 4).setValues([['判定', '下限', '上限', '状態']]).setFontWeight('bold').setBackground('#cfd8dc');
  sh.getRange(25, 1, VALUE_JUDGE.length, 4).setValues(VALUE_JUDGE);

  // ── スキルレベル定義（A29〜C35）
  sh.getRange('A29').setValue('■ スキルレベル定義（1〜5）').setFontWeight('bold');
  sh.getRange(30, 1, 1, 3).setValues([['レベル', '言葉', '意味']]).setFontWeight('bold').setBackground('#cfd8dc');
  sh.getRange(31, 1, SKILL_LEVELS.length, 3).setValues(SKILL_LEVELS);

  // ── 社内固定単価（A37〜B41）
  sh.getRange('A37').setValue('■ 社内固定単価（現場稼働控除用・円／日）').setFontWeight('bold');
  sh.getRange(38, 1, 1, 2).setValues([['稼働内容', '単価']]).setFontWeight('bold').setBackground('#cfd8dc');
  sh.getRange(39, 1, DAILY_RATES.length, 2).setValues(DAILY_RATES);
  sh.getRange(39, 2, DAILY_RATES.length, 1).setNumberFormat('#,##0');

  // ── 運用情報（A43〜B44）
  sh.getRange('A43').setValue('■ 運用情報').setFontWeight('bold');
  sh.getRange('A44').setValue('対象期');
  sh.getRange('B44').setValue(periodLabel).setFontWeight('bold').setBackground('#fff9c4');

  // ── 昇降給マトリクス（F1〜I9）
  sh.getRange('F1').setValue('■ 昇降給マトリクス（粗利判定×バリュー判定→最終判定）').setFontWeight('bold');
  sh.getRange(2, 6).setValue('粗利判定＼バリュー');
  sh.getRange(2, 7, 1, 3).setValues([PAY_MATRIX_COLS]);
  sh.getRange(2, 6, 1, 4).setFontWeight('bold').setBackground('#cfd8dc');
  for (var i = 0; i < PAY_MATRIX_ROWS.length; i++) {
    sh.getRange(3 + i, 6).setValue(PAY_MATRIX_ROWS[i]).setFontWeight('bold').setBackground('#eceff1');
  }
  sh.getRange(3, 7, PAY_MATRIX.length, 3).setValues(PAY_MATRIX).setHorizontalAlignment('center');

  // ── 改定額テーブル（F11〜H19）
  sh.getRange('F11').setValue('■ 改定額（最終判定×等級グループ・円）').setFontWeight('bold');
  sh.getRange(12, 6).setValue('最終判定＼グループ');
  sh.getRange(12, 7, 1, 2).setValues([RAISE_COLS]);
  sh.getRange(12, 6, 1, 3).setFontWeight('bold').setBackground('#cfd8dc');
  for (var j = 0; j < RAISE_ROWS.length; j++) {
    sh.getRange(13 + j, 6).setValue(RAISE_ROWS[j]).setFontWeight('bold').setBackground('#eceff1');
  }
  sh.getRange(13, 7, RAISE_TABLE.length, 2).setValues(RAISE_TABLE).setNumberFormat('+#,##0;-#,##0;±0');

  // ── 項目一覧（K列：バリュー、M列：スキル）
  sh.getRange('K1').setValue('■ バリュー項目').setFontWeight('bold');
  sh.getRange(2, 11, VALUE_ITEMS.length, 1).setValues(VALUE_ITEMS.map(function (v) { return [v]; }));
  sh.getRange('M1').setValue('■ スキル項目').setFontWeight('bold');
  sh.getRange(2, 13, SKILL_ITEMS.length, 1).setValues(SKILL_ITEMS.map(function (v) { return [v]; }));

  sh.setColumnWidths(1, 8, 130);
  sh.getRange('A1').setNote('このシートはGAS（constants.gs）から書き出された参照用です。制度変更はコードを直して再生成してください。');
}

/* ============================ 名簿シート ============================ */
function buildRosterSheet_(ss) {
  var sh = resetSheet_(ss, SHEETS.ROSTER);
  var headers = ['社員ID', '氏名', '区分', '等級', 'B年次', '現在月給', 'ユニット', 'ユニット長', '入社日', 'パートナー報酬形態', '等級キー(自動)'];
  writeHeader_(sh, headers);

  var n = MAX_MEMBERS;
  // 等級キー（K列=11）のみ数式
  var colFns = [];
  for (var c = 0; c < headers.length; c++) colFns.push(null);
  colFns[10] = function (r) {
    return '=IF($D' + r + '="","",IF($D' + r + '="B","B"&IF($E' + r + '="",1,MIN($E' + r + ',3)),$D' + r + '))';
  };
  applyGrid_(sh, 2, n, colFns);

  // 入力規則
  setListValidation_(sh, 3, n, ROSTER_KUBUN);
  setListValidation_(sh, 4, n, ROSTER_GRADES);
  setListValidation_(sh, 5, n, ROSTER_BYEAR);
  setListValidation_(sh, 10, n, PARTNER_PLANS);
  sh.getRange(2, 6, n, 1).setNumberFormat('#,##0');   // 現在月給
  sh.getRange(2, 9, n, 1).setNumberFormat('yyyy/mm/dd'); // 入社日
  sh.setColumnWidths(1, headers.length, 110);
  sh.getRange('A1').setNote('社員IDをキーに全シートが連動します。1行=1人。上から詰めて入力してください。');
}

/* ============================ 粗利入力シート ============================ */
function buildGpSheet_(ss) {
  var sh = resetSheet_(ss, SHEETS.GP);
  var headers = ['社員ID', '氏名', '等級', '等級キー', '半期基準粗利', '実績半期粗利', '粗利倍率', '粗利判定', '経理確定'];
  writeHeader_(sh, headers);
  var n = MAX_MEMBERS;
  var R = SHEETS.ROSTER;
  var colFns = [
    function (r) { return mirror_(R, 'A', r); },                         // 社員ID
    function (r) { return mirror_(R, 'B', r); },                         // 氏名
    function (r) { return mirror_(R, 'D', r); },                         // 等級
    function (r) { return mirror_(R, 'K', r); },                         // 等級キー
    function (r) { return '=IF($D' + r + '="","",IFERROR(VLOOKUP($D' + r + ',' + REF.GRADE_TBL + ',6,FALSE),""))'; }, // 基準粗利
    null,                                                                 // 実績（入力）
    function (r) { return '=IF(OR($E' + r + '="",$F' + r + '=""),"",$F' + r + '/$E' + r + ')'; }, // 倍率
    function (r) { return '=IF($G' + r + '="","",LOOKUP($G' + r + ',' + REF.GPJ_LOWER + ',' + REF.GPJ_JUDGE + '))'; }, // 判定
    null,                                                                 // 経理確定（チェック）
  ];
  applyGrid_(sh, 2, n, colFns);
  sh.getRange(2, 5, n, 2).setNumberFormat('#,##0');
  sh.getRange(2, 7, n, 1).setNumberFormat('0.00"倍"');
  sh.getRange(2, 9, n, 1).insertCheckboxes();
  sh.setColumnWidths(1, headers.length, 120);
  sh.getRange('A1').setNote('実績半期粗利は担当者が入力し、経理が確定します。B等級は固定目標が本体で、倍率は参考値です。');
}

/* ============================ バリュー評価シート ============================ */
function buildValueSheet_(ss) {
  var sh = resetSheet_(ss, SHEETS.VALUE);
  var headers = ['社員ID', '氏名'];
  VALUE_ITEMS.forEach(function (v) { headers.push('自己_' + v); });
  headers.push('自己_合計');
  VALUE_ITEMS.forEach(function (v) { headers.push('長_' + v); });
  headers.push('長_合計', '確定_合計', 'バリュー判定');
  writeHeader_(sh, headers);

  var n = MAX_MEMBERS, R = SHEETS.ROSTER;
  // 列位置：自己 C(3)..K(11), 自己合計 L(12), 長 M(13)..U(21), 長合計 V(22), 確定 W(23), 判定 X(24)
  var colFns = [];
  for (var c = 0; c < headers.length; c++) colFns.push(null);
  colFns[0] = function (r) { return mirror_(R, 'A', r); };
  colFns[1] = function (r) { return mirror_(R, 'B', r); };
  colFns[11] = function (r) { return '=IF(COUNT(C' + r + ':K' + r + ')=0,"",SUM(C' + r + ':K' + r + '))'; };
  colFns[21] = function (r) { return '=IF(COUNT(M' + r + ':U' + r + ')=0,"",SUM(M' + r + ':U' + r + '))'; };
  colFns[22] = function (r) { return '=IF($A' + r + '="","",IF($V' + r + '<>"",$V' + r + ',$L' + r + '))'; };
  colFns[23] = function (r) { return '=IF($W' + r + '="","",IFS($W' + r + '>=23,"高い",$W' + r + '>=17,"標準",TRUE,"要改善"))'; };
  applyGrid_(sh, 2, n, colFns);

  // 0〜3の入力規則（自己 C:K, 長 M:U）
  setListValidation_(sh, 3, n, [0, 1, 2, 3], 9);
  setListValidation_(sh, 13, n, [0, 1, 2, 3], 9);
  sh.setColumnWidths(1, headers.length, 78);
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 90);
  sh.getRange('A1').setNote('9項目×3点＝27点満点。自己→ユニット長の順で入力。確定合計は長合計を既定で採用（会議で上書き可）。');
}

/* ============================ スキル評価シート ============================ */
function buildSkillSheet_(ss) {
  var sh = resetSheet_(ss, SHEETS.SKILL);
  var headers = ['社員ID', '氏名'];
  SKILL_ITEMS.forEach(function (v) { headers.push('自己_' + v); });
  SKILL_ITEMS.forEach(function (v) { headers.push('長_' + v); });
  SKILL_ITEMS.forEach(function (v) { headers.push('確定_' + v); });
  headers.push('全項目3以上', '制作・現場4以上', '1・2残存');
  writeHeader_(sh, headers);

  var n = MAX_MEMBERS, R = SHEETS.ROSTER;
  var ni = SKILL_ITEMS.length; // 10
  // 自己 C(3).. , 長 (3+ni)=13.., 確定 (3+2ni)=23..(22+ni)=32, フラグ 33,34,35
  var confStart = 3 + 2 * ni;  // 23
  var confEnd = confStart + ni - 1; // 32
  var fieldStart = confEnd - SKILL_FIELD_COUNT + 1; // 29 (AC)
  var colFns = [];
  for (var c = 0; c < headers.length; c++) colFns.push(null);
  colFns[0] = function (r) { return mirror_(R, 'A', r); };
  colFns[1] = function (r) { return mirror_(R, 'B', r); };
  // 確定_各項目（長が空なら自己）
  for (var j = 0; j < ni; j++) {
    (function (j) {
      var selfL = colLetter_(3 + j);
      var chofL = colLetter_(3 + ni + j);
      var confC = confStart + j;
      colFns[confC - 1] = function (r) {
        return '=IF($A' + r + '="","",IF(' + chofL + r + '<>"",' + chofL + r + ',' + selfL + r + '))';
      };
    })(j);
  }
  var cs = colLetter_(confStart), ce = colLetter_(confEnd);
  var fs = colLetter_(fieldStart);
  colFns[confEnd] = function (r) {     // 全項目3以上 (col confEnd+1 = 33)
    return '=IF(COUNT(' + cs + r + ':' + ce + r + ')<' + ni + ',"",IF(MIN(' + cs + r + ':' + ce + r + ')>=3,"○","×"))';
  };
  colFns[confEnd + 1] = function (r) { // 制作・現場4以上 (34)
    return '=IF(COUNT(' + cs + r + ':' + ce + r + ')<' + ni + ',"",IF(MAX(' + fs + r + ':' + ce + r + ')>=4,"○","×"))';
  };
  colFns[confEnd + 2] = function (r) { // 1・2残存 (35)
    return '=IF(COUNT(' + cs + r + ':' + ce + r + ')<' + ni + ',"",IF(COUNTIF(' + cs + r + ':' + ce + r + ',"<=2")>0,"有","無"))';
  };
  applyGrid_(sh, 2, n, colFns);

  setListValidation_(sh, 3, n, [1, 2, 3, 4, 5], ni);          // 自己
  setListValidation_(sh, 3 + ni, n, [1, 2, 3, 4, 5], ni);     // 長
  sh.setColumnWidths(1, headers.length, 74);
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 90);
  sh.getRange('A1').setNote('1〜5の5段階×10項目。昇給額には使わず、昇格判断と次のアサイン材料に使う。');
}

/* ============================ 評価結果サマリ ============================ */
function buildSummarySheet_(ss) {
  var sh = resetSheet_(ss, SHEETS.SUMMARY);
  var headers = ['社員ID', '氏名', '等級', '等級グループ', '粗利倍率', '粗利判定', 'バリュー判定', '最終判定', '改定額', '現在月給', '改定後月給', '昇格候補', '降格候補', '会議メモ'];
  writeHeader_(sh, headers);
  var n = MAX_MEMBERS, R = SHEETS.ROSTER;
  var colFns = [
    function (r) { return mirror_(R, 'A', r); },                                  // 社員ID
    function (r) { return mirror_(R, 'B', r); },                                  // 氏名
    function (r) { return mirror_(R, 'D', r); },                                  // 等級
    function (r) { return '=IF($A' + r + '="","",IFERROR(VLOOKUP(\'' + R + '\'!$K' + r + ',' + REF.GRADE_TBL + ',8,FALSE),""))'; }, // グループ
    function (r) { return '=IF($A' + r + '="","",\'' + SHEETS.GP + '\'!$G' + r + ')'; },     // 倍率
    function (r) { return '=IF($A' + r + '="","",\'' + SHEETS.GP + '\'!$H' + r + ')'; },     // 粗利判定
    function (r) { return '=IF($A' + r + '="","",\'' + SHEETS.VALUE + '\'!$X' + r + ')'; },  // バリュー判定
    function (r) { return '=IF(OR($F' + r + '="",$G' + r + '=""),"",IFERROR(INDEX(' + REF.MATRIX_VAL + ',MATCH($F' + r + ',' + REF.MATRIX_ROW + ',0),MATCH($G' + r + ',' + REF.MATRIX_COL + ',0)),""))'; }, // 最終判定
    function (r) { return '=IF(OR($H' + r + '="",$D' + r + '=""),"",IFERROR(INDEX(' + REF.RAISE_VAL + ',MATCH($H' + r + ',' + REF.RAISE_ROW + ',0),MATCH($D' + r + ',' + REF.RAISE_COL + ',0)),""))'; }, // 改定額
    function (r) { return '=IF($A' + r + '="","",\'' + R + '\'!$F' + r + ')'; },             // 現在月給
    function (r) { return '=IF(OR($I' + r + '="",$J' + r + '=""),"",$J' + r + '+$I' + r + ')'; }, // 改定後月給
    function (r) { return '=IF($A' + r + '="","",\'' + SHEETS.PROMOTE + '\'!$J' + r + ')'; }, // 昇格候補
    function (r) { return '=IF($A' + r + '="","",\'' + SHEETS.PROMOTE + '\'!$O' + r + ')'; }, // 降格候補
    null,                                                                                     // 会議メモ
  ];
  applyGrid_(sh, 2, n, colFns);
  sh.getRange(2, 5, n, 1).setNumberFormat('0.00"倍"');
  sh.getRange(2, 9, n, 1).setNumberFormat('+#,##0;-#,##0;±0');
  sh.getRange(2, 10, n, 2).setNumberFormat('#,##0');
  sh.setColumnWidths(1, headers.length, 100);
  sh.setColumnWidth(14, 240);
  sh.getRange('A1').setNote('1人1行の会議用サマリ。判定は自動。会議メモに確定理由・調整を記録します。');
}

/* ============================ 昇格・降格判定 ============================ */
function buildPromoteSheet_(ss) {
  var sh = resetSheet_(ss, SHEETS.PROMOTE);
  var headers = ['社員ID', '氏名', '現等級',
    '粗利連続(A以上3期/S2期)', 'バリュー標準以上', 'スキル全項目3以上', 'スキル強み(制作・現場4以上)',
    '本人意思', 'アサイン・支援体制', '昇格候補可否',
    '降格:粗利D以下2連続', '降格:バリュー要改善2連続', '降格:スキル1・2が2連続', '回復見られず', '降格候補可否'];
  writeHeader_(sh, headers);
  var n = MAX_MEMBERS, R = SHEETS.ROSTER;
  var colFns = [];
  for (var c = 0; c < headers.length; c++) colFns.push(null);
  colFns[0] = function (r) { return mirror_(R, 'A', r); };
  colFns[1] = function (r) { return mirror_(R, 'B', r); };
  colFns[2] = function (r) { return mirror_(R, 'D', r); };
  // D(4)=GAS, E(5)=数式, F(6)=数式, G(7)=数式
  colFns[4] = function (r) { return '=IF($A' + r + '="","",IF(OR(\'' + SHEETS.SUMMARY + '\'!$G' + r + '="標準",\'' + SHEETS.SUMMARY + '\'!$G' + r + '="高い"),"○","×"))'; };
  colFns[5] = function (r) { return '=IF($A' + r + '="","",\'' + SHEETS.SKILL + '\'!$AG' + r + ')'; };
  colFns[6] = function (r) { return '=IF($A' + r + '="","",\'' + SHEETS.SKILL + '\'!$AH' + r + ')'; };
  // H(8),I(9),N(14)=入力。J,K,L,M,O=GAS
  applyGrid_(sh, 2, n, colFns);

  setListValidation_(sh, 8, n, ['○', '×', '—']);   // 本人意思
  setListValidation_(sh, 14, n, ['○', '×']);        // 回復見られず
  sh.setColumnWidths(1, headers.length, 130);
  sh.getRange('A1').setNote('連続条件（粗利/降格判定）はメニュー「昇格・降格を判定」でGASが履歴から計算します。本人意思・支援体制・回復は手入力。可否は最終的に評価会議で判断。');
}

/* ============================ 履歴 ============================ */
function buildHistorySheet_(ss) {
  var sh = resetSheet_(ss, SHEETS.HISTORY);
  var headers = ['期', '社員ID', '氏名', '等級', '粗利倍率', '粗利判定', 'バリュー判定', 'スキル最低レベル', '最終判定', '改定額', '改定後月給', '昇格・降格'];
  writeHeader_(sh, headers);
  sh.setColumnWidths(1, headers.length, 110);
  sh.getRange('A1').setNote('確定時にメニュー「当期を確定して履歴へ」で1人1行追記されます。連続条件の判定に使用。古い→新しい順に並ぶ前提です。');
}

/* ============================ ユーティリティ ============================ */
/** 名簿の同じ行を参照する数式（空なら空） */
function mirror_(sheetName, col, r) {
  return "=IF('" + sheetName + "'!$" + col + r + '="","",\'' + sheetName + '\'!$' + col + r + ')';
}

/**
 * 列に対しリスト入力規則を設定。
 * startCol から repeat 列分（既定1列）、行 2..(numRows+1) に適用。
 */
function setListValidation_(sheet, startCol, numRows, list, repeat) {
  repeat = repeat || 1;
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  sheet.getRange(2, startCol, numRows, repeat).setDataValidation(rule);
}
