function hello() {
  SpreadsheetApp.getUi().alert('こんにちは、Apps Script！');
}

function writeCells() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A1').setValue('税抜価格');
  sheet.getRange('B1').setValue(1200);
  sheet.getRange('A2').setValue('税込価格');
}
