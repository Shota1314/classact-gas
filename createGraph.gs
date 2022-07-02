function createGraphMain(){
  var arr = shape();
  inputShape(arr);
}

function　shape(){
  var arr = getData();

  //不要なデータを削除
  arr.splice(2,1);
  arr.splice(2,1);
  arr.splice(3,1);
  arr.splice(4,1);
  arr.pop();
  arr.pop();

  return arr;

}

function inputShape(arr){

  var name = arr[0];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);

  var arrNum = arr.length;
  arrNum = arrNum - 2;
  var h = 1;
  var i = 2;
  var k = 1;
  var cellPosition = [h, i];

  while(true){
    var contract = sheet.getRange(cellPosition[0], cellPosition[1]);
    if(contract.isBlank()){
      sheet.getRange(cellPosition[0], cellPosition[1]).setValue(arr[k]);
      while(k <= arrNum){
        k++;
        h++;
        cellPosition[0] = h;
        sheet.getRange(cellPosition[0], cellPosition[1]).setValue(arr[k]);
      }
      break;
    }else{
      i++;
      cellPosition[1] = i;
    }
  }

  }