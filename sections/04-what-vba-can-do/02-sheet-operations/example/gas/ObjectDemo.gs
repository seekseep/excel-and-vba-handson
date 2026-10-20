function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('デモ')
    .addItem('ブックの情報を出す', 'showObjects')
    .addItem('表を整える', 'decorateSales')
    .addItem('顧客別に分ける', 'splitByCustomer')
    .addItem('増えたシートを消す', 'deleteDemoSheets')
    .addItem('バックアップを保存', 'saveBackup')
    .addToUi();
}

function showObjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log('スプレッドシートの名前: ' + ss.getName());
  console.log('シートの数: ' + ss.getSheets().length);
  console.log('1 枚目のシート: ' + ss.getSheets()[0].getName());
  console.log('売上!B2 の値: ' + ss.getSheetByName('売上').getRange('B2').getValue());
}

function decorateSales() {
  const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('売上');
  const lastRow = salesSheet.getLastRow();

  salesSheet.getRange('A1:F1').setBackground('#1f4e79').setFontColor('#ffffff');
  salesSheet.getRange('A1:F' + lastRow).setBorder(true, true, true, true, true, true);

  for (let i = 2; i <= lastRow; i++) {
    const cell = salesSheet.getRange(i, 6);
    if (cell.getValue() >= 5000) {
      cell.setFontColor('#c00000').setFontWeight('bold');
    }
  }
}

function splitByCustomer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('売上');
  const lastRow = salesSheet.getLastRow();

  for (let c = 5; c <= 9; c++) {
    const customerName = salesSheet.getRange(c, 8).getValue();
    const customerSheet = ss.insertSheet(customerName, ss.getSheets().length);
    salesSheet.getRange('A1:F1').copyTo(customerSheet.getRange('A1'));

    let outRow = 2;
    for (let i = 2; i <= lastRow; i++) {
      if (salesSheet.getRange(i, 2).getValue() === customerName) {
        salesSheet.getRange('A' + i + ':F' + i).copyTo(customerSheet.getRange('A' + outRow));
        outRow = outRow + 1;
      }
    }
    customerSheet.autoResizeColumns(1, 6);
  }
  ss.setActiveSheet(salesSheet);
}

function deleteDemoSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const sheet of ss.getSheets()) {
    if (sheet.getName() !== '売上') {
      ss.deleteSheet(sheet);
    }
  }
}

function saveBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const copy = DriveApp.getFileById(ss.getId()).makeCopy('売上_バックアップ_' + stamp);
  SpreadsheetApp.getUi().alert(copy.getName() + ' をドライブに保存しました');
}
