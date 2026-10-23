function calcAmounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  for (let i = 2; i <= 31; i++) {
    const amount = sheet.getRange(i, 4).getValue() * sheet.getRange(i, 5).getValue();
    sheet.getRange(i, 6).setValue(amount);
  }
}
