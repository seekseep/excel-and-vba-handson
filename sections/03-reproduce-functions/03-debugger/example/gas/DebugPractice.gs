// 3-3 デバッグの練習（Apps Script 版）
// この関数は売上合計を J2 に書くが、関数（I2）の結果と合わない。
// ブレークポイントとデバッガーで、どこがおかしいかを探そう。

function sumTotalWithBug() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  let total = 0;
  for (let i = 3; i <= lastRow; i++) {
    total = total + sheet.getRange(i, 6).getValue();
  }
  sheet.getRange('J2').setValue(total);
}
