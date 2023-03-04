/**
* NX_Company
* For mission that the boss gives to me. Google Spreadsheet and App Script were used.
* Copyright (c)Nextersystems, Inc.
*
* If you want to use or are interested in this code, plz email to us js.lee@nextersystems.co.kr :)
*
* Company homepage : https://www.nextersystems.co.kr/
/

const FileId = "파일 ID"

//Data SpreadSheet file id 갖고 오기
function svld_id(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var togetfolder = ss.getSheetByName("Work sheet") // ID 작성된 대상 sheetname
  var id = togetfolder.getRange('C2').getValues(); // 작성된 행열
  return id
}

//Reviewer list 추출
function get_Reviewer_list(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Work sheet');
  let Rsers = []
  vals = sheet.getRange('E2:E').getValues();
  for (let i = 0 ; i < vals.length ; i++ ){
    if ( vals[i][0] == '')
    {
      break;
    }
    Rsers.push(vals[i][0]);
  }
  console.info(Rsers);
  return Rsers
}

//Labeler list 추출
function get_labeler_list(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Work Sheet');
  let Rsers = []
  vals = sheet.getRange('J2:J').getValues();
  for (let i = 0 ; i < vals.length ; i++ ){
    if ( vals[i][0] == '')
    {
      break;
    }
    Rsers.push(vals[i][0]);
  }
  //console.info(Rsers.toString().toLowerCase());
  return Rsers
}

//UI button성생성
function onOpen() {
  SpreadsheetApp.getUi().createMenu('실행(Do)')
      .addItem('labeler 수량 새로고침(Labeler Output)', 'fun_labeler')
      .addItem('Reviewer 수량 새로고침(Reviewer Output)', 'fun_Reviewer')
      .addItem('인원 새로고침(labeler list)','write_lb')
      .addItem('인원 새로고침(reviewer list)','write_Rv')
      .addItem('file save','file_saving')
      .addToUi();
}

//Reviewer Output 산출
function fun_Reviewer(){
  refresh_Rv();
  write_RV_totalre();
  write_RV_totalcom();
}

//Labeler Input and Output 산출
function fun_labeler(){
  refresh_lb();
  write_lb_totalin();
  write_lb_totalout();
}

function write_lb(){
  var labelerlist = get_labeler_list();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('LB of IN&OUT');

  var targetsheet = ss.getSheetByName('this month');

  for(var i=0;i<get_labeler_list().length;i++){
    sheet.getRange(5+2*i,2).setValue(labelerlist[i].toString().toLowerCase()) //레이블러 리스트
  }
}

function write_lb_totalin(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sumsheet = ss.getSheetByName('LB of IN&OUT')
  var totalnum = 0

  for(var i=0;i<get_labeler_list().length;i++){
    var num = sumsheet.getRange(5+i*2,4).getValue()
    totalnum = Number(totalnum) + Number(num)
  }
  sumsheet.getRange('D1').setValue(totalnum);
}

function refresh_lb(){
  var tot = 0;
  var list = get_labeler_list();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.openById(svld_id());
  var sheets = sheet.getSheets();

  for (var k = 0; k < list.length; k++) {
    for (d = 0; d < 31; d++) {
      var num = sheet.getSheetByName(list[k].toString().toLowerCase()).getRange(d + 7, 2).getValue();
      ss.getSheetByName('LB of IN&OUT').getRange(6 + (k * 2), 5 + d).setValue(num);
    }

    var lbnum = sheet.getSheetByName(list[k]).getRange('B4').getValue();
    ss.getSheetByName('LB of IN&OUT').getRange(6 + (k * 2), 4).setValue(Number(lbnum));
  }
}

function write_lb_totalout(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sumsheet = ss.getSheetByName('LB of IN&OUT')
  var totalnum = 0

  for(var i=0;i<get_labeler_list().length;i++){
    var num = sumsheet.getRange(6+i*2,4).getValue()
    totalnum = Number(totalnum) + Number(num)
  }
  sumsheet.getRange('D2').setValue(totalnum);
}

function write_Rv(){
  var Reviewerlist = get_Reviewer_list();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RV of IN&OUT');

  for(var i=0;i<Reviewerlist.length;i++){
    sheet.getRange(5+2*i,2).setValue(Reviewerlist[i])
    console.info(Reviewerlist[i])
  }
}

function refresh_Rv(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var Reviewerlist = get_Reviewer_list(); //list of Reviewer
  
  var sheet = SpreadsheetApp.openById(svld_id());
  var sheets = sheet.getSheets();

  for (var k=0 ; k<Reviewerlist.length ; k++){
    for(d=0 ; d<31 ; d++){
      var num = sheet.getSheetByName(Reviewerlist[k]).getRange(d+7,12).getValue();
      ss.getSheetByName('RV of IN&OUT').getRange(6+(k*2),5+d).setValue(Number(num))
    }
    var Rvnum = sheet.getSheetByName(Reviewerlist[k]).getRange('L4').getValue();
    ss.getSheetByName('RV of IN&OUT').getRange(6+(k*2),4).setValue(Number(Rvnum))
  }

  for (var k=0 ; k<Reviewerlist.length ; k++){
    for(d=0 ; d<31 ; d++){
      var num = sheet.getSheetByName(Reviewerlist[k]).getRange(d+7,11).getValue();
      ss.getSheetByName('RV of IN&OUT').getRange(5+(k*2),5+d).setValue(Number(num))
    }
    var Rvnum = sheet.getSheetByName(Reviewerlist[k]).getRange('K4').getValue();
    ss.getSheetByName('RV of IN&OUT').getRange(5+(k*2),4).setValue(Number(Rvnum))
  }
}

function write_RV_totalre(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sumsheet = ss.getSheetByName('RV of IN&OUT')
  var totalnum = 0

  for(var i=0;i<get_Reviewer_list().length;i++){
    var num = sumsheet.getRange(5+i*2,4).getValue()
    totalnum = Number(totalnum) + Number(num)
  }
  sumsheet.getRange('D1').setValue(totalnum);
}

function write_RV_totalcom(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sumsheet = ss.getSheetByName('RV of IN&OUT')
  var totalnum = 0

  for(var i=0;i<get_Reviewer_list().length;i++){
    var num = sumsheet.getRange(6+i*2,4).getValue()
    totalnum = Number(totalnum) + Number(num)
  }
  sumsheet.getRange('D2').setValue(totalnum);
}

//장파일 저장
function naming_bar()
{
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
      '파일 이름을 입력하세요.(Please enter a file name)',
      'Year-Month:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    return text
  }
}

function exportSheet_(filename) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var newFile = DriveApp.createFile(ss.getBlob().setName(filename));
  
  newFile.moveTo(DriveApp.getFolderById(FileId)); 
  return newFile.getId();
}

function file_saving()
{
  name = naming_bar();
  fileId = exportSheet_(name);

}
