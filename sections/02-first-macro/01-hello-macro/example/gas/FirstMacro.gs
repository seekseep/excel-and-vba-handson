function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('マクロ')
    .addItem('あいさつする', 'hello')
    .addToUi();
}

function hello() {
  SpreadsheetApp.getUi().alert('こんにちは、Apps Script！');
}
