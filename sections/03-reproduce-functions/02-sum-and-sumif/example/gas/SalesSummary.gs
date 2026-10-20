function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('売上集計')
    .addItem('金額を計算', 'calcAmounts')
    .addItem('売上合計を出す', 'sumTotal')
    .addItem('顧客別に集計', 'sumByCustomer')
    .addToUi();
}

function calcAmounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  for (let i = 2; i <= lastRow; i++) {
    const amount = sheet.getRange(i, 4).getValue() * sheet.getRange(i, 5).getValue();
    sheet.getRange(i, 6).setValue(amount);
  }
}

function sumTotal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  let total = 0;
  for (let i = 2; i <= lastRow; i++) {
    total = total + sheet.getRange(i, 6).getValue();
  }
  sheet.getRange('J2').setValue(total);
}

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
