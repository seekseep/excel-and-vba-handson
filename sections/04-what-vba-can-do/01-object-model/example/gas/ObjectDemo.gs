// 4-1 Excel は部品の入れ子（Apps Script 版）
// 結果は実行ログ（console.log）で見る。

// スプレッドシート → シート → セル は入れ子になっている
function showObjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log('スプレッドシートの名前: ' + ss.getName());
  console.log('シートの数: ' + ss.getSheets().length);
  console.log('1 枚目のシート: ' + ss.getSheets()[0].getName());
  console.log('売上!B2 の値: ' + ss.getSheetByName('売上').getRange('B2').getValue());
}
