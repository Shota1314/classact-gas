//このgsファイルで一番最初に呼び出されるfunction
function getDataMain(){
  buttonClick();
  console.log("-----getDataMainが呼び出されました。処理を開始します-----")
  console.log("-----getDataを呼び出します-----")
  var arr = getData();
  console.log("-----getDataから処理が戻ってきました-----")
  console.log("-----getCellを呼び出します-----")
  var cellPosition = getCell(arr);
  console.log("-----getCellから処理が戻ってきました-----")
  console.log("-----putDataを呼び出します-----")
  putData(cellPosition, arr);
  console.log("-----putDataから処理が戻ってきました-----")
  console.log("-----getDataMainの処理が正常終了しました-----");
}

//処理実行前の確認画面出力。誤実行防止の為。
function buttonClick() {
  let msg = Browser.msgBox('確認メッセージ','実行します。よろしいですか？実行しない場合、画面上部に現れるキャンセルを押下するか、ページをリロードして下さい。',Browser.Buttons.OK);
  if (msg == 'ok') {
    Browser.msgBox('OKが押されました。処理を実行します。');
  }
}

//リーダ業務時間計算シートから指定の数値を取得するfunction
function getData(){
  console.log("-----getDataが呼び出されました。処理を開始します-----")
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("リーダ業務時間計算シート ");

  var leaderDateMem = sheet.getRange("D1:D3").getValues();
  var leaderDateMem = Array.prototype.concat.apply([], leaderDateMem);
  leaderDateMem[1] = Utilities.formatDate(leaderDateMem[1], 'JST', 'yyyy-MM');
  console.log("【リーダ：対象年月：管理人数】を取得しました。");
  console.log(leaderDateMem);

  var weektask = sheet.getRange("D7:E9").getValues();
  var weektask = Array.prototype.concat.apply([], weektask);
  console.log("【1人当たりの週報返信：週報返信合計：1人当たりの1on1：1on1合計：1人当たりのHRMOS:HRMOS合計】を取得しました。");
  console.log(weektask);

  var meeting = sheet.getRange("E13").getValues();
  var meeting = Array.prototype.concat.apply([], meeting);
  console.log("【リーダ定例】を取得しました。");
  console.log(meeting);

  var others = sheet.getRange("E17:E21").getValues();
  var others = Array.prototype.concat.apply([], others);
  console.log("【GLOBIS/オーディオセミナー：シフト表管理：リーダタスク：個別帰社日：その他リーダタスク】を取得しました。");
  console.log(others);

  var result = sheet.getRange("D26:D27").getValues();
  var result = Array.prototype.concat.apply([], result);
  console.log("【1人当たりの対応時間：合計】を取得しました。");
  console.log(result);

  //取得したデータを配列に成型
  var arr = leaderDateMem.concat(weektask,meeting,others,result);
  console.log("取得した全てのデータは以下となります。");
  console.log(arr);
  console.log("-----getDataが正常終了しました。getDataMainに処理を戻します-----")
  return arr;
}


function getCell(arr){  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("1人当たりの対応時間");
  var leaderName = arr[0];
  var date = arr[1];
  i = 3;
  j = 2;
  
  //リーダ業務時間計算シートの"年月"データを元にシートの行番号を特定する処理
  console.log("行番号を取得します。")
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
  console.log("行番号を取得しました。行番号は：" + i + "です。")

　//リーダ業務時間計算シートの"リーダ名"データを元にシートの列番号を特定する処理
  console.log("列番号を取得します。")
  while(true){
    var rowCon = sheet.getRange(2, j).getValue();
    if(leaderName == rowCon){
      rowCon = j;
      break;
    }else{
      j++;
    }
  }
  console.log("列番号を取得しました。列番号は：" + j + "です。")

  var cellPosition = [i, j];
  console.log("-----getCellが正常終了しました。getDataMainに処理を戻します-----")
  return cellPosition;
}

function putData(cellPosition, arr){
  var content = ["管理人数","1人当たりの週報返信時間","週報返信時間合計","1人当たりの1on1時間","1on1時間合計","1人当たりのHRMOS月締め合計","HRMOS月締め合計","リーダ定例合計","GLOBIS_オーディオセミナー合計","シフト管理表合計","リーダミッション合計","個別帰社日合計","その他リーダタスク合計","1人当たりの対応時間","1か月合計時間"];
    
  arr.shift(); arr.shift();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contentNum = content.length;
  for(var k=0; k<contentNum; k++){
    var sheet = ss.getSheetByName(content[k]);  
    sheet.getRange(cellPosition[0], cellPosition[1]).setValue(arr[k]);
  }
}
