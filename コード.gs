//GASを使ったテスト結果の転記
function main(){
  planSheetOpen();
  var file = getFilesByNameRgeExp();
//記録に使うJSONファイル
  Logger.log(file);
  resultArray = readJson(file);
//  Logger.log(resultArray);
  resultSheetOpen();
  recResult(resultArray);
}

//テスト計画シートを開いて、JSONテスト結果から利用可能にしておく
//テスト実施時刻は、テスト計画に記載の時刻を記録する。MagicPodの実時間ではない。
function planSheetOpen() {
  var today　= new Date();
  var formatDate = Utilities.formatDate(today, "JST","yyyy/M/d");
  let spreadSheetByActive = SpreadsheetApp.getActive();
  let sheetByActive = spreadSheetByActive.getActiveSheet();
  let sheetByName   = spreadSheetByActive.getSheetByName("テスト計画");
}

function resultSheetOpen(){
  let spreadSheetByActive = SpreadsheetApp.getActive();
  let sheetByName   = spreadSheetByActive.getSheetByName("テスト結果");
}

//magicpod-analyzerで取得したjsonファイルは、自動テスト記録スプシと同じGoogleドライブに置く
function getFilesByNameRgeExp(){
  // フォルダの指定
  const folderId= "ここにGoogleDriveのフォルダID"
 //フォルダ内のすべてのファイルを取得
  const folder = DriveApp.getFolderById(folderId);

  //ファイルタイプを指定
  const name = /magicpod.*/; //ファイル名が「magicpod…～」となっているファイルが対象
  const files = folder.getFiles(); // すべてのファイルを取得

  //各ファイルごとに出力
  while(files.hasNext()){
    let file = files.next();

   // ファイル名が「name」でしている正規表現にマッチすれば実行
    if(file.getName().match(name)){
      let fileName = file.getName(); // ファイル名
      let fileId = file.getId(); // ファイルID
      let fileURL = file.getUrl(); // ファイルURL
//      Logger.log([fileName,fileId,fileURL]);
      return fileName
    }
  }
}

function readJson(file_name) {
  var fileIT = DriveApp.getFilesByName(file_name).next();
  var textdata = fileIT.getBlob().getDataAsString('utf8');
  var jobj = JSON.parse(textdata);

  var retArray = [];

//テスト名と実施日とテスト結果のJSONを受け取る
  const flatJson = jobj.map((e) => testStatus(e))
//配列に格納する
  const jsonResults = Object.entries(flatJson);
  const testResults = Object.values(flatJson);
//  Logger.log(testResults);

//JSONファイルの結果は全て記録対象
  var suffix = 0;
  jsonResults.forEach(function(jsonResults){
    //テスト設定名
    var workflowName = Object.values(jsonResults)[1]["workflowName"];
    Logger.log(workflowName);
    //テスト実施日
//JSTに再計算
    var createdAtUTC = Object.values(jsonResults)[1]["createdAt"];
    var createdAt = calcJst(createdAtUTC).slice(0, 10);
//UTCで返ってくるので（2022/7/1現在）そのままでは使えない。
//    var createdAt = (Object.values(jsonResults)[1]["createdAt"]).slice(0, 10);
    Logger.log(createdAt);
    //テスト実施時刻、テスト計画から取得する
    var timedata = searchTestCase(createdAt, workflowName);
    Logger.log(timedata);
    //テスト結果
    var status = Object.values(jsonResults)[1]["status"];
    Logger.log(status);
    retArray.push([workflowName, createdAt, timedata, status]);
    suffix = suffix + 1;
  });
//  }
  return retArray;
}

//magicpod-analyzerから取得したJSONをテスト名と実施日とテスト結果だけにする
const testStatus = (obj) => {
  const result = {};
  for (const key in obj){
    const value = obj[key];
    if (typeof value === "object"){

    }else if ((key == "createdAt") || (key == "workflowName") || (key == "status")){
      result[key] = value;
    }
  }
  return result;
};

const flattenObj = (obj) => {
  const result = {};

    // rootにあるkeyごとに処理
  for (const key in obj) {

    // 値を取得
    const value = obj[key];

    // value がObjectだった場合はflattenObj(自分自身)を呼び出して処理
    if (typeof value === "object") {
      const flatObj = flattenObj(value);

      // key と subkey を結合して root の key とする
      for (const subKey in flatObj) {
        result[`${key}.${subKey}`] = flatObj[subKey];
      }
    } else {
      result[key] = value;
    }
  }
  return result;
};

//テスト計画シートから、実行時刻を取得
function searchTestCase(createdAt, workflowName){
  //名前付き範囲を作成しておき、テスト結果の曜日から範囲を選んで使う。
　var arr_day = new Array('rangeSunday', 'rangeMonday', 'rangeTuesday', 'rangeWednesday', 'rangeThursday', 'rangeFriday', 'rangeSaturday');

　var bk = SpreadsheetApp.getActiveSpreadsheet();
　var sh = bk.getSheetByName("テスト計画");
　//var rng = sh.getActiveCell();

　var day_num = new Date(createdAt).getDay();
  var rangeName = arr_day[day_num]
  Logger.log(rangeName);
//選んだ範囲を配列に格納する
　var rngArray = bk.getRangeByName(rangeName).getValues();
//  Logger.log(rngArray);
//  Logger.log(workflowName);
  var timeData = getTestTime(sh, rngArray, workflowName);
//  Logger.log(timeData);
  return timeData;
}

function getTestTime(sh, rngArray, workflowName){
  var pattern = '/' + workflowName + '/';
    
  for(var i = 0; i < rngArray.length; i++) 
  {
    if(rngArray[i][0].indexOf(workflowName) !== -1)
    {
      var timeRow = i + 1;
      var timeDataCell = 'A' + timeRow;
      var timeData = sh.getRange(timeDataCell).getValue(); 
    }
  }
  return timeData;
}

function calcJst(date){
  if (date === '' ) {
    return ''
  } else {
    const d = new Date(date);
    var format = 'yyyy-MM-dd HH:mm:ss';
    var retval = Utilities.formatDate(d, 'JST', format);
//    Logger.log(retval);
//    Logger.log(date + '->' + retval + '(' + timeZone +')');
    return retval;
  }
}

function recResult(resultArray){
  var bk = SpreadsheetApp.getActiveSpreadsheet();
  let lastRow = bk.getLastRow();
  let lastCol = 49;
　var sh = bk.getSheetByName("テスト結果").activate();
  var testDay = sh.getRange(1, 1, lastRow).getValues();
  
//  Logger.log(testDay.length);
  for(var suffix = 0; suffix < resultArray.length; suffix++){
  //A列からテスト実施日を探す
    for(var tD = 0; tD < testDay.length; tD++){
//        Logger.log(testDay[tD][0]);
//        Logger.log(resultArray[suffix][1]);
      if(testDay[tD][0].indexOf(resultArray[suffix][1]) !== -1){
        var recRow = tD + 1;
      }
    }
//    Logger.log(recRow);
    //1行目からテスト実施時刻を探す
    for(var col = 1; col <= lastCol; col++){
//        Logger.log(sh.getRange(1, col).getValue());
//        Logger.log(resultArray[suffix][2]);
      if(sh.getRange(1, col).getValue() === resultArray[suffix][2]){
        var recCol = col;
      }
    }
//    Logger.log(recCol);
    var testStatus = resultArray[suffix][3];
    if(testStatus == 'SUCCESS'){
      sh.getRange(recRow, recCol).activate();
      sh.getRange(recRow, recCol).setValue('success');
      sh.getActiveRangeList().setBackground('#0000ff');  
    }
    if(testStatus == 'FAILURE'){
      sh.getRange(recRow, recCol).activate();
      sh.getRange(recRow, recCol).setValue('failed');
      sh.getActiveRangeList().setBackground('#ff0000');  
    }
  } 

}