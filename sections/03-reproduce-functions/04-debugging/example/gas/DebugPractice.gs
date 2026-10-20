function sumTotalWithBug() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = sheet.getLastRow();
  let total = 0;
  for (let i = 3; i <= lastRow; i++) {
    total = total + sheet.getRange(i, 6).getValue();
  }
  sheet.getRange('J2').setValue(total);
}
