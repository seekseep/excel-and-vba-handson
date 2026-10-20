// 2-2 セルと変数（Apps Script 版）
// 2-1 のコードに続けて書き、上から順に実行する。

// メッセージを出すだけのマクロ
function hello() {
  SpreadsheetApp.getUi().alert('こんにちは、Apps Script！');
}

// セルに書き込む
function writeCells() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A1').setValue('税抜価格');
  sheet.getRange('B1').setValue(1200);
  sheet.getRange('A2').setValue('税込価格');
}

// セルから読んで、計算して、セルに書く
function calcTaxIncluded() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const price = sheet.getRange('B1').getValue();
  const taxIncluded = Math.floor(price * 1.1);
  sheet.getRange('B2').setValue(taxIncluded);
}
