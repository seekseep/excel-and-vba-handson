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

  // 売上からこの顧客の行だけを明細に転記する
  const lastSalesRow = salesSheet.getLastRow();
  let lineRow = 12;
  let total = 0;
  for (let i = 2; i <= lastSalesRow; i++) {
    if (salesSheet.getRange(i, 2).getValue() === customerName) {
      invoiceSheet.getRange(lineRow, 1).setValue(salesSheet.getRange(i, 1).getValue());
      invoiceSheet.getRange(lineRow, 2).setValue(salesSheet.getRange(i, 3).getValue());
      invoiceSheet.getRange(lineRow, 3).setValue(salesSheet.getRange(i, 4).getValue());
      invoiceSheet.getRange(lineRow, 4).setValue(salesSheet.getRange(i, 5).getValue());
      invoiceSheet.getRange(lineRow, 5).setValue(salesSheet.getRange(i, 6).getValue());
      total = total + salesSheet.getRange(i, 6).getValue();
      lineRow = lineRow + 1;
    }
  }
  invoiceSheet.getRange('E33').setValue(total);
  invoiceSheet.getRange('B9').setValue(total);
}
