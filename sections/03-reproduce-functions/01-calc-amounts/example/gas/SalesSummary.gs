// 3-1 金額を計算する（Apps Script 版）
// data/sales.xlsx をスプレッドシートに変換し、その Apps Script エディタに貼り付ける。

// 金額 = 単価 × 数量 を F 列の全行に書く（関数なら F2 に =D2*E2 を入れて下までコピー）
function calcAmounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  for (let i = 2; i <= lastRow; i++) {
    const amount = sheet.getRange(i, 4).getValue() * sheet.getRange(i, 5).getValue();
    sheet.getRange(i, 6).setValue(amount);
  }
}
