//GAS���g�����e�X�g���ʂ̓]�L
function main(){
  planSheetOpen();
  var file = getFilesByNameRgeExp();
//�L�^�Ɏg��JSON�t�@�C��
  Logger.log(file);
  resultArray = readJson(file);
//  Logger.log(resultArray);
  resultSheetOpen();
  recResult(resultArray);
}

//�e�X�g�v��V�[�g���J���āAJSON�e�X�g���ʂ��痘�p�\�ɂ��Ă���
//�e�X�g���{�����́A�e�X�g�v��ɋL�ڂ̎������L�^����BMagicPod�̎����Ԃł͂Ȃ��B
function planSheetOpen() {
  var today�@= new Date();
  var formatDate = Utilities.formatDate(today, "JST","yyyy/M/d");
  let spreadSheetByActive = SpreadsheetApp.getActive();
  let sheetByActive = spreadSheetByActive.getActiveSheet();
  let sheetByName   = spreadSheetByActive.getSheetByName("�e�X�g�v��");
}

function resultSheetOpen(){
  let spreadSheetByActive = SpreadsheetApp.getActive();
  let sheetByName   = spreadSheetByActive.getSheetByName("�e�X�g����");
}

//magicpod-analyzer�Ŏ擾����json�t�@�C���́A�����e�X�g�L�^�X�v�V�Ɠ���Google�h���C�u�ɒu��
function getFilesByNameRgeExp(){
  // �t�H���_�̎w��
  const folderId= "������GoogleDrive�̃t�H���_ID"
 //�t�H���_���̂��ׂẴt�@�C�����擾
  const folder = DriveApp.getFolderById(folderId);

  //�t�@�C���^�C�v���w��
  const name = /magicpod.*/; //�t�@�C�������umagicpod�c�`�v�ƂȂ��Ă���t�@�C�����Ώ�
  const files = folder.getFiles(); // ���ׂẴt�@�C�����擾

  //�e�t�@�C�����Ƃɏo��
  while(files.hasNext()){
    let file = files.next();

   // �t�@�C�������uname�v�ł��Ă��鐳�K�\���Ƀ}�b�`����Ύ��s
    if(file.getName().match(name)){
      let fileName = file.getName(); // �t�@�C����
      let fileId = file.getId(); // �t�@�C��ID
      let fileURL = file.getUrl(); // �t�@�C��URL
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

//�e�X�g���Ǝ��{���ƃe�X�g���ʂ�JSON���󂯎��
  const flatJson = jobj.map((e) => testStatus(e))
//�z��Ɋi�[����
  const jsonResults = Object.entries(flatJson);
  const testResults = Object.values(flatJson);
//  Logger.log(testResults);

//JSON�t�@�C���̌��ʂ͑S�ċL�^�Ώ�
  var suffix = 0;
  jsonResults.forEach(function(jsonResults){
    //�e�X�g�ݒ薼
    var workflowName = Object.values(jsonResults)[1]["workflowName"];
    Logger.log(workflowName);
    //�e�X�g���{��
//JST�ɍČv�Z
    var createdAtUTC = Object.values(jsonResults)[1]["createdAt"];
    var createdAt = calcJst(createdAtUTC).slice(0, 10);
//UTC�ŕԂ��Ă���̂Łi2022/7/1���݁j���̂܂܂ł͎g���Ȃ��B
//    var createdAt = (Object.values(jsonResults)[1]["createdAt"]).slice(0, 10);
    Logger.log(createdAt);
    //�e�X�g���{�����A�e�X�g�v�悩��擾����
    var timedata = searchTestCase(createdAt, workflowName);
    Logger.log(timedata);
    //�e�X�g����
    var status = Object.values(jsonResults)[1]["status"];
    Logger.log(status);
    retArray.push([workflowName, createdAt, timedata, status]);
    suffix = suffix + 1;
  });
//  }
  return retArray;
}

//magicpod-analyzer����擾����JSON���e�X�g���Ǝ��{���ƃe�X�g���ʂ����ɂ���
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

    // root�ɂ���key���Ƃɏ���
  for (const key in obj) {

    // �l���擾
    const value = obj[key];

    // value ��Object�������ꍇ��flattenObj(�������g)���Ăяo���ď���
    if (typeof value === "object") {
      const flatObj = flattenObj(value);

      // key �� subkey ���������� root �� key �Ƃ���
      for (const subKey in flatObj) {
        result[`${key}.${subKey}`] = flatObj[subKey];
      }
    } else {
      result[key] = value;
    }
  }
  return result;
};

//�e�X�g�v��V�[�g����A���s�������擾
function searchTestCase(createdAt, workflowName){
  //���O�t���͈͂��쐬���Ă����A�e�X�g���ʂ̗j������͈͂�I��Ŏg���B
�@var arr_day = new Array('rangeSunday', 'rangeMonday', 'rangeTuesday', 'rangeWednesday', 'rangeThursday', 'rangeFriday', 'rangeSaturday');

�@var bk = SpreadsheetApp.getActiveSpreadsheet();
�@var sh = bk.getSheetByName("�e�X�g�v��");
�@//var rng = sh.getActiveCell();

�@var day_num = new Date(createdAt).getDay();
  var rangeName = arr_day[day_num]
  Logger.log(rangeName);
//�I�񂾔͈͂�z��Ɋi�[����
�@var rngArray = bk.getRangeByName(rangeName).getValues();
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
�@var sh = bk.getSheetByName("�e�X�g����").activate();
  var testDay = sh.getRange(1, 1, lastRow).getValues();
  
//  Logger.log(testDay.length);
  for(var suffix = 0; suffix < resultArray.length; suffix++){
  //A�񂩂�e�X�g���{����T��
    for(var tD = 0; tD < testDay.length; tD++){
//        Logger.log(testDay[tD][0]);
//        Logger.log(resultArray[suffix][1]);
      if(testDay[tD][0].indexOf(resultArray[suffix][1]) !== -1){
        var recRow = tD + 1;
      }
    }
//    Logger.log(recRow);
    //1�s�ڂ���e�X�g���{������T��
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