/** 現在のシート名を表示する
 * @customfunction
 * 
 */
function getSheetName() {
  //const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //const sheetName = sheet.getName();  return SpreadsheetApp.getActiveSheet().getName();
  //return sheetName;

  return SpreadsheetApp.getActiveSheet().getName();
}


/** スプレッドシート内のシート名一覧を返す
 * @customfunction
 * 
 */
function allSheetNames(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(sheet => sheet.getName());
  return sheets;
}


/** DBの名前と役割の表示をいい感じにするやつ
 * @customfunction
 * 
 */
function replaceAndRemoveDuplicates() {
  // スプレッドシートとシートのIDを指定
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('データベース'); // シート名を適宜変更
  var range = sheet.getRange('B1:B100'); // 変換したい範囲を指定

  // 指定範囲のデータを取得
  var values = range.getValues();
  
  // 値を格納するセットを用意
  var uniqueValues = new Set();
  
  // 各セルの値を変換し、セットに追加
  for (var i = 0; i < values.length; i++) {
    var cellValue = values[i][0];
    // 'wp_' の後に数字が続く部分を 'wp_[site_id]_' に置換
    var newValue = cellValue.replace(/wp_\d+_/g, 'wp_[site_id]_');
    uniqueValues.add(newValue);
  }
  
  // 重複を排除した値をスプレッドシートに反映
  var uniqueArray = Array.from(uniqueValues).map(value => [value]);
  var uniqueRange = sheet.getRange(1, 5, uniqueArray.length, 1);
  uniqueRange.setValues(uniqueArray);
  
  // 余った行をクリア
  if (uniqueArray.length < values.length) {
    var clearRange = sheet.getRange(uniqueArray.length + 1, 1, values.length - uniqueArray.length, 1);
    clearRange.clearContent();
  }
}
