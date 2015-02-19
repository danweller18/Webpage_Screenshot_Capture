/*****************************************************************
'          Title:  Webpage Screenshot Capture V1.0
'   Published by:  www.funbutlearn.com
'      Copyright:  © 2014 FunButLearn
'****************************************************************/



var GOOGLE_DRIVE_FOLDER_NAME = "Web page captures";



function onOpen() {
  //This code will run on opening the sheet, It will add the menu bar at the top
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [
  {name: "Start capturing", functionName: "startTriggers"},
  {name: "Stop capturing", functionName: "stopTriggers"},
  {name: "Verify now", functionName: "verifyCapturing"}
  ];
  sheet.addMenu("Capture Webpage", menu);
  sheet.toast("You can also start the app from menu bar.", "Message", 7);
}

function startTriggers(){
    var sheet = getSheet();
    var triggers = ScriptApp.getProjectTriggers();
  if(triggers.length>0){
    Browser.msgBox("Message", "Application is already started !", Browser.Buttons.OK);
    return false;
  }
    Logger.log(sheet.getRange("D9").getValue());
    switch(sheet.getRange("D9").getValue()){
       case "Every weekday":
        ScriptApp.newTrigger("captureWebPage").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).create();
        ScriptApp.newTrigger("captureWebPage").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).create();
        ScriptApp.newTrigger("captureWebPage").timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).create();
        ScriptApp.newTrigger("captureWebPage").timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).create();
        ScriptApp.newTrigger("captureWebPage").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).create();
        break;
      case "Once in a day":
        ScriptApp.newTrigger("captureWebPage").timeBased().everyDays(1).create();
        break;
      case "Once in a week":
        ScriptApp.newTrigger("captureWebPage").timeBased().everyDays(7).create();
        break;
      case "Once in a month":
        ScriptApp.newTrigger("captureWebPage").timeBased().everyDays(30).create();
        break;
    }
    Browser.msgBox("Message", "Application is started and ready to capture the web page.", Browser.Buttons.OK);
}

function stopTriggers(){
    var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
    Browser.msgBox("Message", "Application has been stopped.", Browser.Buttons.OK);
}


function captureWebPage() {
  var sheet = getSheet();
  createFolderIfNotExist();
  var url = "http://api.page2images.com/restfullink?p2i_key="+sheet.getRange("D3").getValue()+"&p2i_url="+
    encodeURIComponent(sheet.getRange("D6").getValue());

  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText();
  var data1 = JSON.parse(json);
  var data = data1.image_url;
  var image = UrlFetchApp.fetch(data);
  var folder = DocsList.getFolder(GOOGLE_DRIVE_FOLDER_NAME);
  var result = folder.createFile(image);
  var time = new Date().toLocaleString();
  result.rename(time+" ("+sheet.getRange("D6").getValue()+").png");
  sheet.getRange("A"+getLastRow()).setValue(time);
  sheet.getRange("D12").setValue(sheet.getRange("D12").getValue()+1);
  var id = result.getId();
  var file = DriveApp.getFileById(id);
  var date = Utilities.formatDate(new Date(), "GMT-5", "MMM dd, yyyy");
  var email = sheet.getRange("D15").getValue();
  var emailsub = sheet.getRange("D18").getValue();
  var emailbody = sheet.getRange("D21").getValue();
  MailApp.sendEmail(email, emailsub, emailbody + date, {
     name: emailsub + date,
     attachments: [file.getAs(MimeType.PNG)]
 });
}

function verifyCapturing(){
  captureWebPage();
  Browser.msgBox("Message","Web page has been captured. Visit your Google Drive to view the captured image.",Browser.Buttons.OK);
}

function getLastRow(){
  var sheet = getSheet();
  var lastRow = sheet.getLastRow()+1;
  if(sheet.getLastRow()<=30){
    var i = 1;
    while(sheet.getRange("A"+i).getValue()!=""){
      i++;
    }
    lastRow = i;
  }
  return lastRow;
}

function getSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
}

function createFolderIfNotExist(){
  var sheet   = SpreadsheetApp.getActiveSheet();
  var folders = DocsList.getAllFolders();
  var exist = false;
  for (var x=0; x<folders.length; x++) {
    if (folders[x].getName() == GOOGLE_DRIVE_FOLDER_NAME) {
      exist = true;
      break;
    }
  }
  if (!exist) {
    DocsList.createFolder(GOOGLE_DRIVE_FOLDER_NAME);
  }
}
