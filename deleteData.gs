function deleteDataMain() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("リーダ業務時間計算シート ");  

  let range = sheet.getRange('D1:D21');
  range.clearContent();

}