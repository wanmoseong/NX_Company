/**
* NX_Company
* For mission that the boss gives to me. Google Spreadsheet and App Script were used.
* Copyright (c)Nextersystems, Inc.
*
* If you want to use or are interested in this code, plz email to us js.lee@nextersystems.co.kr :)
* Thank you.
* Company homepage : https://www.nextersystems.co.kr/
/

//folder ID
function experiment(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var togetfolder = ss.getSheetByName("Search")
  var id = togetfolder.getRange('D7').getValues();
  //console.info(id);
  return id
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('발송 및 기타(Sending and Other)')
      .addItem('전체 보내기(Entire Sending)', 'allSend_')
      .addItem('개인 보내기(Individual Sending)', 'singleSend_')
      .addItem('통계 수집(Collect statistic from SVLD)', 'setStaticSheet')
      .addItem('수량 갱신(Calculate Quantity)', 'update_')
      .addItem('인보이스 코드 생성(Invoicecode maker)','make_invoicecode')
      .addToUi();
}

function get_Reviewer(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Search');
  let Rsers = []
  vals = sheet.getRange('I2:I').getValues();
  for (let i = 0 ; i < vals.length ; i++ ){
    if ( vals[i][0] == '')
    {
      break;
    }
    Rsers.push(vals[i][0]);
  }
  //console.info(Rsers);
  return Rsers
}


function getUsers(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Search');
  let users = []
  vals = sheet.getRange('A2:A').getValues();
  for (let i = 0 ; i < vals.length ; i++ ){
    if ( vals[i][0] == '')
    {
      break;
    }
    users.push(vals[i][0]);
  }
  //console.info(users);
  return users
}

function make_invoicecode(){
  var list = getUsers();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Search');
  var date1 = sheet.getRange('D6').getValue()
  var date2 = sheet.getRange('E6').getValue()

  for (var i = 0 ; i<list.length; i++){
    //console.info('NX-LD-'+date1+date2+list[i][0].toUpperCase()+list[i][1].toUpperCase()+list[i][2].toUpperCase())
    sheet.getRange(2+i,2).setValue('NX-LD-'+date1+date2+list[i][0].toUpperCase()+list[i][1].toUpperCase()+list[i][2].toUpperCase())
  }
}

function isExistName(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() == sheetName) {
      return true
    }
  }
  return false
}

function getEmailAddr(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  return sheet.getRange('F10').getValue();
}

function exportSheet_(sheetName) {
  var date = new Date();
  var strDate = date.getFullYear()+"-"+(date.getMonth()+1)+"-"+date.getDate();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() !== sheetName) {
      sheets[i].hideSheet()
    }
  }
  var newFile = DriveApp.createFile(ss.getBlob().setName(sheetName+"_"+strDate));
  for (var i = 0; i < sheets.length; i++) {
    sheets[i].showSheet()
  }
  newFile.moveTo(DriveApp.getFolderById(experiment())); 
  return newFile.getId();

}

function test_t(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Search');
  var dateof = sheet.getRange('D11:E12').getValues();
  var day = dateof[0][0] +"/" + dateof[0][1];
  console.info(day);
}

function sendEmail_(emailAddr, attachFileId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Search');
  var dateof = sheet.getRange('D11:E12').getValues();
  var day = dateof[0][0] +"/" + dateof[0][1];
  var subjectText = sheet.getRange('D13').getValue();
  var bodyText = sheet.getRange('D16').getValue();
  var nameText = sheet.getRange('D17').getValue();
  console.info("sendEmail_:"+emailAddr+","+attachFileId+", day="+day);
  var file = DriveApp.getFileById(attachFileId);
  var message = {
    to: emailAddr,
    subject: subjectText, // B1 셀의 내용을 subject에 넣음
    body: "\n" + bodyText.replace(/\n/g, '\n\n'),
    name: nameText, // C1 셀의 내용을 name에 넣음
    attachments: [file.getAs(MimeType.PDF)]
  }
  MailApp.sendEmail(message);
}

function promptSendEmail_()
{
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
      'Invoice Mail을 발송',
      'ID를 넣으세요.:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    return text
  }
  return ''
}

function allSend_()
{
  getUsers().forEach( (value, index, array) =>{
    fileId = exportSheet_(value);
    emailAddr = getEmailAddr(value)
    sendEmail_(emailAddr,fileId);
  });
}

function singleSend_()
{
  name = promptSendEmail_();
  if ( name == '' ){
    return;
  }
  if ( !isExistName(name) ){
    return;
  }

  fileId = exportSheet_(name);
  emailAddr = getEmailAddr(name)
  sendEmail_(emailAddr,fileId);

}

function allViewStatistic(sheetId){
  var sheet = SpreadsheetApp.openById(sheetId);
  var out = new Array()
  var sheets = sheet.getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    var wsname = sheets[i].getName();
    if (wsname == 'label' || wsname == 'draw' || wsname == 'tempwork' || wsname == 'dashboard'){
      continue;
    }
    var values = sheets[i].getRange('B4:K4').getValues();
    var rateError = 0;
    if (  values[0][0] > 0 ){
      rateError = values[0][9]*100.0 / values[0][0];
    }
    out.push( [wsname , values[0][0] , values[0][1] , values[0][9], rateError] );
  }
  //console.info(out);
  return out 
}

function setStaticSheet(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('Static');
  let sheetId = ws.getRange('B1').getValue();
  console.info("sheet id="+sheetId);

  var lastRow = ws.getMaxRows();
  ws.insertRowAfter(lastRow); 
  var lastRow = ws.getMaxRows();
  if( 4 < lastRow){
    console.info(lastRow);
    ws.deleteRows(3, lastRow-3 );
  }

  statistic = allViewStatistic(sheetId);

  for (i = 0; i < statistic.length; i++) {

    ws.appendRow(
      statistic[i]
    )
  }   
}

function update_(){
  refresh_maureen()
  refresh_labeler()
}

function refresh_maureen(){
  var tot = 0;
  var list = getUsers();
  list.shift();
  //console.info(list);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('Static');
  var sheetId = ws.getRange('B1').getValue();
  var sheet = SpreadsheetApp.openById(sheetId);

  //OUT 작성
  for (var k=0 ; k<list.length ; k++){
    console.info(list[k].toLowerCase())
    var lbnum = sheet.getSheetByName(list[k].toLowerCase()).getRange('B4').getValue();
    tot = tot + Number(lbnum);
    console.info(lbnum);
  }

  var targetSheet = ss.getSheetByName('Maureen');
  targetSheet.getRange('J20').setValue(tot);
}

function refresh_labeler(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('Static');
  let sheetId = ws.getRange('B1').getValue();
  statistic = allViewStatistic(sheetId);
    var sheet = ss.getSheetByName('Static');
    var lbvalue
    var num = 0
    var lbname
    var first_comparename = ss.getSheets();
    for(var d = 0; d<first_comparename.length; d++){
      var comparename = first_comparename[d].getName();
      for(var i = 0; i<statistic.length ; i++){
        lbname = sheet.getRange(3+i,1).getValue();
        if (lbname == get_Reviewer()){
          continue;
        }
        else if (lbname == comparename.toLowerCase()){
          console.info(comparename)
          var targetSheet = ss.getSheetByName(comparename);
          lbvalue = sheet.getRange(3+i,2).getValue();
          targetSheet.getRange('J21').setValue(lbvalue);
          //console.info(lbname,'실행')
      }
    }
}
}
