function calcAmounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const amount = sheet.getRange(2, 4).getValue() * sheet.getRange(2, 5).getValue();
  sheet.getRange(2, 6).setValue(amount);
}
