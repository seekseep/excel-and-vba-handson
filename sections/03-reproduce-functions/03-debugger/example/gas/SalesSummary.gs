// 3-3 デバッガーで原因を探す（Apps Script 版）
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

// 売上合計を J2 に書く（関数なら =SUM(F2:F31)）
function sumTotal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  let total = 0;
  for (let i = 2; i <= lastRow; i++) {
    total = total + sheet.getRange(i, 6).getValue();
  }
  sheet.getRange('J2').setValue(total);
}

// H5 の顧客の合計を J5 に書く（関数なら =SUMIF(B2:B31, H5, F2:F31)）
function sumByCustomer() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  const customerName = sheet.getRange('H5').getValue();
  let total = 0;
  for (let i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, 2).getValue() === customerName) {
      total = total + sheet.getRange(i, 6).getValue();
    }
  }
  sheet.getRange('J5').setValue(total);
}

// チャレンジ: H5〜H9 の全顧客の合計を J5〜J9 に書く（for の中に for）
function sumAllCustomers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  for (let c = 5; c <= 9; c++) {
    let total = 0;
    for (let i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 2).getValue() === sheet.getRange(c, 8).getValue()) {
        total = total + sheet.getRange(i, 6).getValue();
      }
    }
    sheet.getRange(c, 10).setValue(total);
  }
}
