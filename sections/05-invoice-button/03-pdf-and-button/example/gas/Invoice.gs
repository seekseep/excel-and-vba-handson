// 5-3 PDF にしてボタンから実行する（Apps Script 版）
// スプレッドシートを開くとメニュー「請求書」が出る。図形描画のボタンに createInvoices を割り当ててもよい。
// 請求書テンプレートのセル番地（A5 宛名 / B9 請求金額 / 12 行目から明細 / E33 合計）に合わせて書いている。

// スプレッドシートを開いたときに自動で動き、メニューを足す
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('請求書')
    .addItem('請求書を作る', 'createInvoices')
    .addItem('請求書シートを消す', 'deleteInvoices')
    .addToUi();
}

// 顧客ごとに請求書シートを作り、PDF にする
function createInvoices() {
  // 何度押しても同じ結果になるよう、前回の請求書シートを消してから作る
  deleteInvoices();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('売上');
  const customerSheet = ss.getSheetByName('顧客');
  const lastSalesRow = salesSheet.getLastRow();
  const lastCustomerRow = customerSheet.getLastRow();

  for (let c = 2; c <= lastCustomerRow; c++) {
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

    // 書き込みを確定させてから、スプレッドシートと同じフォルダに PDF を保存する
    SpreadsheetApp.flush();
    savePdf_(ss, invoiceSheet);
  }

  ss.setActiveSheet(customerSheet);
  SpreadsheetApp.getUi().alert((lastCustomerRow - 1) + ' 件の請求書を作りました');
}

// 「請求書_」で始まるシートを消す
function deleteInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const sheet of ss.getSheets()) {
    if (sheet.getName().startsWith('請求書_')) {
      ss.deleteSheet(sheet);
    }
  }
}

// シート 1 枚を PDF にしてドライブに保存する。
// Excel の ExportAsFixedFormat にあたる命令が無いので、スプレッドシートの「ダウンロード」と同じ URL を呼ぶ。
// 名前の末尾の _ は「実行する関数の一覧に出さない」という Apps Script の決まり。
function savePdf_(ss, sheet) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&gid=' + sheet.getSheetId();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  });
  const folder = DriveApp.getFileById(ss.getId()).getParents().next();
  folder.createFile(response.getBlob().setName(sheet.getName() + '.pdf'));
}
