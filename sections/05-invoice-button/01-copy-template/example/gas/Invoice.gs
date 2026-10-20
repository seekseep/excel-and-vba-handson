function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('請求書')
    .addItem('請求書を作る', 'createInvoices')
    .addToUi();
}

function createInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('売上');
  const customerSheet = ss.getSheetByName('顧客');

  const c = 2;
  const customerName = customerSheet.getRange(c, 1).getValue();

  // テンプレートを末尾にコピーして名前を付ける
  const invoiceSheet = ss.getSheetByName('請求書テンプレート').copyTo(ss);
  invoiceSheet.setName('請求書_' + customerName);

  // 宛名と発行日
  invoiceSheet.getRange('A5').setValue(customerName + ' 御中');
  invoiceSheet.getRange('A6').setValue(customerSheet.getRange(c, 2).getValue() + ' ' + customerSheet.getRange(c, 3).getValue());
  invoiceSheet.getRange('A7').setValue(customerSheet.getRange(c, 4).getValue());
  invoiceSheet.getRange('E3').setValue(new Date());
}
