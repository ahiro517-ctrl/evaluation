/**
 * interview.gs
 * 個人面談シート（1人1枚の縦型フォーム）。
 * 設定マスタを参照して自分の判定まで自己完結で計算する。
 * 名簿から1人ずつシートを複製生成する。会議用サマリへはコピペで取り込む前提
 *  （シート間のライブ連携は組まない）。
 */

/** 個人面談シートのテンプレートを生成 */
function buildInterviewTemplate_(ss) {
  var sh = resetSheet_(ss, SHEETS.INTERVIEW_TPL);
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 150);
  sh.setColumnWidth(4, 150);

  // タイトル
  sh.getRange('A1:D1').merge().setValue('個人面談シート')
    .setFontSize(16).setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 基本情報
  setLabel_(sh, 'A2', '対象期');           sh.getRange('B2').setFormula("='設定'!$B$44");
  setLabel_(sh, 'A3', '氏名');
  setLabel_(sh, 'A4', '社員ID');
  setLabel_(sh, 'A5', '区分');
  setLabel_(sh, 'A6', '等級');             setLabel_(sh, 'C6', 'B年次');
  setLabel_(sh, 'A7', 'ユニット');         setLabel_(sh, 'C7', 'ユニット長');
  setLabel_(sh, 'A8', '現在月給');

  // 補助セル（等級キー・等級グループ）→ 列E:Fは非表示
  sh.getRange('E1').setValue('等級キー'); sh.getRange('E2').setValue('等級グループ');
  sh.getRange('F1').setFormula('=IF($B$6="","",IF($B$6="B","B"&IF($D$6="",1,MIN($D$6,3)),$B$6))');
  sh.getRange('F2').setFormula('=IFERROR(VLOOKUP($F$1,' + REF.GRADE_TBL + ',8,FALSE),"")');

  // ── 粗利評価
  section_(sh, 9, '■ 粗利評価');
  setLabel_(sh, 'A10', '半期基準粗利'); sh.getRange('B10').setFormula('=IF($F$1="","",IFERROR(VLOOKUP($F$1,' + REF.GRADE_TBL + ',6,FALSE),""))');
  setInput_(sh, 'A11', '実績半期粗利', 'B11');
  setLabel_(sh, 'A12', '粗利倍率');     sh.getRange('B12').setFormula('=IF(OR($B$10="",$B$11=""),"",$B$11/$B$10)');
  setLabel_(sh, 'A13', '粗利判定');     sh.getRange('B13').setFormula('=IF($B$12="","",LOOKUP($B$12,' + REF.GPJ_LOWER + ',' + REF.GPJ_JUDGE + '))');
  sh.getRange('B10').setNumberFormat('#,##0');
  sh.getRange('B11').setNumberFormat('#,##0');
  sh.getRange('B12').setNumberFormat('0.00"倍"');
  sh.getRange('A11').setNote('B等級は固定目標が本体。倍率は規模感をつかむ参考値です。');

  // ── バリュー評価
  section_(sh, 15, '■ バリュー評価（各3点・9項目・27点満点）');
  sh.getRange('A16:C16').setValues([['項目', '自己', 'ユニット長']]).setFontWeight('bold').setBackground('#cfd8dc');
  var vStart = 17;
  sh.getRange(vStart, 1, VALUE_ITEMS.length, 1).setValues(VALUE_ITEMS.map(function (v) { return [v]; }));
  var vEnd = vStart + VALUE_ITEMS.length - 1; // 25
  var totalRow = vEnd + 1;       // 26 合計
  var confRow = vEnd + 2;        // 27 確定合計
  var vjRow = vEnd + 3;          // 28 バリュー判定
  setLabel_(sh, 'A' + totalRow, '合計');
  sh.getRange(totalRow, 2).setFormula('=IF(COUNT(B' + vStart + ':B' + vEnd + ')=0,"",SUM(B' + vStart + ':B' + vEnd + '))');
  sh.getRange(totalRow, 3).setFormula('=IF(COUNT(C' + vStart + ':C' + vEnd + ')=0,"",SUM(C' + vStart + ':C' + vEnd + '))');
  setLabel_(sh, 'A' + confRow, '確定合計（既定＝ユニット長／会議で上書き可）');
  sh.getRange(confRow, 2).setFormula('=IF($B$4="","",IF(C' + totalRow + '<>"",C' + totalRow + ',B' + totalRow + '))');
  setLabel_(sh, 'A' + vjRow, 'バリュー判定');
  sh.getRange(vjRow, 2).setFormula('=IF($B$' + confRow + '="","",IFS($B$' + confRow + '>=23,"高い",$B$' + confRow + '>=17,"標準",TRUE,"要改善"))');
  setListRule_(sh.getRange(vStart, 2, VALUE_ITEMS.length, 2), [0, 1, 2, 3]); // 自己・長

  // ── スキル評価
  var sTitle = vjRow + 2;        // 30
  section_(sh, sTitle, '■ スキル評価（1〜5・10項目）※昇給額には使わない');
  var sHead = sTitle + 1;        // 31
  sh.getRange(sHead, 1, 1, 4).setValues([['項目', '自己', 'ユニット長', '確定']]).setFontWeight('bold').setBackground('#cfd8dc');
  var skStart = sHead + 1;       // 32
  sh.getRange(skStart, 1, SKILL_ITEMS.length, 1).setValues(SKILL_ITEMS.map(function (v) { return [v]; }));
  var skEnd = skStart + SKILL_ITEMS.length - 1; // 41
  // 確定列D：長が空なら自己
  var dFns = [];
  for (var i = 0; i < SKILL_ITEMS.length; i++) {
    var r = skStart + i;
    dFns.push(['=IF($B$4="","",IF(C' + r + '<>"",C' + r + ',B' + r + '))']);
  }
  sh.getRange(skStart, 4, SKILL_ITEMS.length, 1).setFormulas(dFns);
  // フラグ
  var fldStart = skEnd - SKILL_FIELD_COUNT + 1; // 制作・現場の先頭行 38
  var f1 = skEnd + 1, f2 = skEnd + 2, f3 = skEnd + 3; // 42,43,44
  setLabel_(sh, 'A' + f1, '全項目3以上');
  sh.getRange(f1, 2).setFormula('=IF(COUNT(D' + skStart + ':D' + skEnd + ')<' + SKILL_ITEMS.length + ',"",IF(MIN(D' + skStart + ':D' + skEnd + ')>=3,"○","×"))');
  setLabel_(sh, 'A' + f2, '制作・現場で4以上の強み');
  sh.getRange(f2, 2).setFormula('=IF(COUNT(D' + skStart + ':D' + skEnd + ')<' + SKILL_ITEMS.length + ',"",IF(MAX(D' + fldStart + ':D' + skEnd + ')>=4,"○","×"))');
  setLabel_(sh, 'A' + f3, '1・2が残っている');
  sh.getRange(f3, 2).setFormula('=IF(COUNT(D' + skStart + ':D' + skEnd + ')<' + SKILL_ITEMS.length + ',"",IF(COUNTIF(D' + skStart + ':D' + skEnd + ',"<=2")>0,"有","無"))');
  setListRule_(sh.getRange(skStart, 2, SKILL_ITEMS.length, 2), [1, 2, 3, 4, 5]); // 自己・長

  // ── 最終判定・処遇
  var pTitle = f3 + 2;           // 46
  section_(sh, pTitle, '■ 最終判定・処遇');
  var finRow = pTitle + 1, raiseRow = pTitle + 2, nowRow = pTitle + 3, newRow = pTitle + 4;
  setLabel_(sh, 'A' + finRow, '最終判定（粗利×バリュー）');
  sh.getRange(finRow, 2).setFormula('=IF(OR($B$13="",$B$' + vjRow + '=""),"",IFERROR(INDEX(' + REF.MATRIX_VAL + ',MATCH($B$13,' + REF.MATRIX_ROW + ',0),MATCH($B$' + vjRow + ',' + REF.MATRIX_COL + ',0)),""))');
  setLabel_(sh, 'A' + raiseRow, '改定額');
  sh.getRange(raiseRow, 2).setFormula('=IF(OR($B$' + finRow + '="",$F$2=""),"",IFERROR(INDEX(' + REF.RAISE_VAL + ',MATCH($B$' + finRow + ',' + REF.RAISE_ROW + ',0),MATCH($F$2,' + REF.RAISE_COL + ',0)),""))');
  setLabel_(sh, 'A' + nowRow, '現在月給');
  sh.getRange(nowRow, 2).setFormula('=$B$8');
  setLabel_(sh, 'A' + newRow, '改定後月給');
  sh.getRange(newRow, 2).setFormula('=IF(OR($B$' + raiseRow + '="",$B$' + nowRow + '=""),"",$B$' + nowRow + '+$B$' + raiseRow + ')');
  sh.getRange(raiseRow, 2).setNumberFormat('+#,##0;-#,##0;±0');
  sh.getRange(nowRow, 2, 2, 1).setNumberFormat('#,##0');

  // ── 昇格・降格メモ
  var gTitle = newRow + 2;       // 53
  section_(sh, gTitle, '■ 昇格・降格メモ（連続条件は履歴を見て評価会議で判断）');
  setInput_(sh, 'A' + (gTitle + 1), '昇格への本人意思', 'B' + (gTitle + 1));
  setListRule_(sh.getRange(gTitle + 1, 2), ['○', '×', '—']);
  setInputWide_(sh, 'A' + (gTitle + 2), '昇格後のアサイン・支援体制', gTitle + 2);
  setInputWide_(sh, 'A' + (gTitle + 3), '降格の懸念・回復状況', gTitle + 3);

  // ── 面談メモ
  var mTitle = gTitle + 5;       // 58
  section_(sh, mTitle, '■ 面談メモ（次の半期に向けて）');
  var memo = ['自分の粗利', '不足している要素', '次に取り組むこと', '伸ばすスキル', '必要な案件経験', '会社・チームに求めるサポート'];
  for (var m = 0; m < memo.length; m++) {
    var rr = mTitle + 1 + m;
    setInputWide_(sh, 'A' + rr, memo[m], rr);
  }

  // ── 転記用（会議シートへコピペ）
  var tTitle = mTitle + 1 + memo.length + 1;
  section_(sh, tTitle, '■ 転記用（値をコピーして会議シートへ貼り付け）');
  var t = tTitle + 1;
  setLabel_(sh, 'A' + t, '粗利実績 →「粗利入力」F列');             sh.getRange(t, 2).setFormula('=$B$11');
  setLabel_(sh, 'A' + (t + 1), 'バリュー自己 →「バリュー評価」C列'); putStrip_(sh, t + 1, 2, vStart, vEnd, 'B');
  setLabel_(sh, 'A' + (t + 2), 'バリュー長 →「バリュー評価」M列');   putStrip_(sh, t + 2, 2, vStart, vEnd, 'C');
  setLabel_(sh, 'A' + (t + 3), 'スキル自己 →「スキル評価」C列');     putStrip_(sh, t + 3, 2, skStart, skEnd, 'B');
  setLabel_(sh, 'A' + (t + 4), 'スキル長 →「スキル評価」M列');       putStrip_(sh, t + 4, 2, skStart, skEnd, 'C');
  sh.getRange(t, 1, 5, 1).setFontColor('#616161').setFontSize(9);

  sh.hideColumns(5, 2); // E:F 補助列を隠す
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('1人1枚の面談用シート。自己→ユニット長の順で入力。判定は自動。会議用サマリへは「転記用」の値をコピペで取り込む。');
  return sh;
}

/** 名簿から1人ずつ面談シートを複製生成（既存はスキップ） */
function generateInterviewSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tpl = ss.getSheetByName(SHEETS.INTERVIEW_TPL);
  if (!tpl) { buildInterviewTemplate_(ss); tpl = ss.getSheetByName(SHEETS.INTERVIEW_TPL); }
  var roster = ss.getSheetByName(SHEETS.ROSTER);
  if (!roster) throw new Error('名簿シートがありません。先に評価シートを生成してください。');

  var data = roster.getRange(2, 1, MAX_MEMBERS, 8).getValues(); // ID..ユニット長
  var made = 0, skipped = 0;
  for (var i = 0; i < data.length; i++) {
    var id = data[i][0];
    if (id === '' || id === null) continue;
    var name = SHEETS.INTERVIEW_PREFIX + id;
    if (ss.getSheetByName(name)) { skipped++; continue; }
    var ns = tpl.copyTo(ss).setName(name);
    ns.showSheet();
    ns.getRange('B3').setValue(data[i][1]); // 氏名
    ns.getRange('B4').setValue(id);         // 社員ID
    ns.getRange('B5').setValue(data[i][2]); // 区分
    ns.getRange('B6').setValue(data[i][3]); // 等級
    ns.getRange('D6').setValue(data[i][4]); // B年次
    ns.getRange('B7').setValue(data[i][6]); // ユニット
    ns.getRange('D7').setValue(data[i][7]); // ユニット長
    ns.getRange('B8').setValue(data[i][5]); // 現在月給
    made++;
  }
  ss.toast('面談シートを ' + made + ' 件作成しました（既存スキップ ' + skipped + ' 件）。', '評価運用', 6);
}

/* ── ヘルパー ── */
function setLabel_(sh, a1, text) {
  sh.getRange(a1).setValue(text).setFontWeight('bold').setBackground('#eceff1');
}
function setInput_(sh, a1Label, text, a1Input) {
  setLabel_(sh, a1Label, text);
  sh.getRange(a1Input).setBackground('#fffde7');
}
/** ラベル(A)＋B:Dを結合した広い入力欄 */
function setInputWide_(sh, a1Label, text, row) {
  setLabel_(sh, a1Label, text);
  sh.getRange(row, 2, 1, 3).merge().setBackground('#fffde7').setWrap(true).setVerticalAlignment('top');
}
function section_(sh, row, text) {
  sh.getRange(row, 1, 1, 4).merge().setValue(text).setFontWeight('bold')
    .setBackground('#455a64').setFontColor('#ffffff');
}
function setListRule_(range, list) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  range.setDataValidation(rule);
}
/** srcCol列の startRow..endRow を、row行の startCol列から横並びに参照 */
function putStrip_(sh, row, startCol, startRow, endRow, srcCol) {
  var fns = [];
  for (var r = startRow; r <= endRow; r++) fns.push('=$' + srcCol + '$' + r);
  sh.getRange(row, startCol, 1, fns.length).setFormulas([fns]);
}
