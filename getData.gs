function getDataMain(){
  var arr = getData();
  logger(arr);
  var cellPosition = getCell(arr);
  putData(cellPosition, arr);
}

function getData() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("リーダ業務時間計算シート ");

  var nameDateMem = sheet.getRange("D1:D3").getValues();
  var nameDateMem = Array.prototype.concat.apply([], nameDateMem);
  nameDateMem[1] = Utilities.formatDate(nameDateMem[1], 'JST', 'yyyy-MM');

  var weekrepoOneoneHrmos = sheet.getRange("D7:E9").getValues();
  var weekrepoOneoneHrmos = Array.prototype.concat.apply([], weekrepoOneoneHrmos);

  var meeting = sheet.getRange("E13").getValues();
  var meeting = Array.prototype.concat.apply([], meeting);

  var others = sheet.getRange("E17:E21").getValues();
  var others = Array.prototype.concat.apply([], others);

  var result = sheet.getRange("D25:D26").getValues();
  var result = Array.prototype.concat.apply([], result);

  //取得したデータを配列に成型
  var arr = nameDateMem.concat(weekrepoOneoneHrmos,meeting,others,result);
  return arr;
}

function logger(arr){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ログ");
  sheet.appendRow(arr);
}

function getCell(arr){  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("1人当たりの対応時間");
  var name = arr[0];
  var date = arr[1];
  i = 3;
  j = 2;
  
  while(true){
    var col = "A" + i;
    var colCon = sheet.getRange(col).getValue();
    colCon = Utilities.formatDate(colCon, 'JST', 'yyyy-MM');
    if(date == colCon){
      colCon = i;
      break;
    }else{
      i++;
    }
  }

  while(true){
    var rowCon = sheet.getRange(2, j).getValue();
    if(name == rowCon){
      rowCon = j;
      break;
    }else{
      j++;
    }
  }

  var cellPosition = [i, j];
  return cellPosition;
}

function putData(cellPosition, arr){
  var content = ["担当者","日付","管理人数","1人当たりの週報返信時間","週報返信時間合計","1人当たりの1on1時間","1on1時間合計","1人当たりのHRMOS月締め合計","HRMOS月締め合計","リーダ定例合計","GLOBIS_オーディオセミナー合計","シフト管理表合計","リーダミッション合計","個別帰社日合計","その他リーダタスク合計","1人当たりの対応時間","1か月合計時間"];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contentNum = content.length;
  for(var k=2; k<contentNum; k++){
    var sheet = ss.getSheetByName(content[k]);  
    sheet.getRange(cellPosition[0], cellPosition[1]).setValue(arr[k]);
  }



}
