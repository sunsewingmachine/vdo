// https://stackoverflow.com/questions/63935779/comparing-values-of-two-different-cells-evaluate-to-false-google-apps-script
// (([Urgency] = "*") AND (ISBLANK([Status])))

// DONOT DELETE
// Note: E3 uses 'https://app.multcloud.com/mc_project/cloud_sync' to sync '' folder from Google to Onedrive
// If google drive is installed in E2 pc, we should stop this service. It is 3rd party service.

var PauseExecution = false;
var IsBulkUpdateDoingNow = false;
const IsExtraDisplayEnabled = false;
const RangeSheetStatus = "B1";
// const RangeSheetStatus = "K1";
const RangeTimeNow = "Db!A3";
const RangeTimeLastExecuted = "Db!A5";
const ScheduleSheetsRangeInDB = "Db!C6:C39";
const sheetWorksDeleted = "WorksDeleted";
var UseSelectedCellAsMonth = false;

const colNoSNo = 1;         // A
const colNoWork = 2;        // B
const colNoStatus = 3;      // C
const colNoRepeat = 4;      // D

const colNoProperTime = 5;  // E
const colNoDue = 6;         // F
const colNoStaff = 7;       // G
const colNoUrgency = 8;     // H

// const colNoExtra1 = 9;     // I - Info - Change this const name to someother or remove, ambugious
const colNoUid = 9;           // I - Info
const colNoReport = 10;       // J -
const colNoExtraDisplay1 = 11;// K - 
const colNoDetails = 12;      // L - Details
const colNoTags = 13;         // M - Tags (show/hide/etc to use in AppSheet)
const colNoEditedTime = 14;         // N - Tags (show/hide/etc to use in AppSheet)
// Leave upto column 18 for future works

const colNoJobReportDay1 =  20; // T
// ==============================================================================

const StartRow = 2;  // 2 
var colTotal = 12;      


// Sno	Job	Status	Repeat	Job time	Postponed	Staff	Urgency	Info	Report	Extra1	Details	Tags
// ==============================================================================

const colNametSNo = "Sno";              // A
const colNameWork = "Job";              // B
const colNameStatus = "Status";         // C
const colNameRepeat = "Repeat";         // D

const colNameProperTime = "Job time";   // E
const colNameDue = "Postponed";         // F
const colNameStaff = "Staff";           // G
const colNameUrgency = "Urgency";       // H

const colNameExtra1 = "Uid";           // I - Info - Change this const name to someother or remove, ambugious
const colNameUid = "Uid";             // I - Info
const colNameReport = "Report";         // J -
const colNameExtraDisplay1 = "Extra1";  // K - 
const colNameDetails = "Details";       // L - Details
const colNameTags = "Tags";             // T - Tags - 20(show/hide/etc to use in AppSheet) 
// ==============================================================================


const WorkNew = "WorkNew";
const WorkUpdate = "Work-Update";
const WorkShowFullField = "Work-ShowFullField";
const WorkHideDoneWorks = "Work-HideDoneWorks";
const WorkHideDoneWorksNow = "Work-HideDoneWorksNow";

const FolderSsmcJobReports = "https://drive.google.com/drive/folders/1pCVXY-HIfDOlZHWWubEvur5eG92JhRhV";
// ==============================================================================


function openUrl1(url) {

  // url =  FolderSsmcJobReports;
  var htmlOutput = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");</script>');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening URL');
}

function openUrl2(url){
  
  // url =  FolderSsmcJobReports;
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}


function BackupThisGoogleSheet(){

  // Folder name: My Drive/Gdrive-Backups/GoogleSheets/Vdo
  // https://drive.google.com/drive/folders/1zRfYM5O7h_tyajKWEKN5jjmBWj4PQJ0W

  var sourceFile = DriveApp.getFileById('1jQ-6typqFmEAk6wtlJKdxlw-yWD7cZo-OuAUsqSrzKI');
  var destFolder = DriveApp.getFolderById('1zRfYM5O7h_tyajKWEKN5jjmBWj4PQJ0W');
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yy")
  var destName = 'Backup of - ' + SpreadsheetApp.getActiveSpreadsheet().getName() + ' (' + date + ')';

  sourceFile.makeCopy(destName, destFolder);
}

function autoHideWorks(){  
  // If bShowAll=true, unhide all rows
  var bShowAll = getShowAll();
  LoggerLog("getShowAll(): " + bShowAll);  

}


function SetTitle(activSpreadsheet) {
  //var dataAry = [['SNo', 'Job','Status','Repeat','Job time','Postponed','Staff','Urgency','A','Report']];
  //var dataAry = [['SNo', 'Job','Status','Repeat','Job time','Postponed','Staff','Urgency','Info','Report','JobDetail']];
  //activSpreadsheet.getRange('A1:K1').setValues(dataAry);

  var dataAry = [[colNametSNo, colNameWork, colNameStatus, colNameRepeat, colNameProperTime, colNameDue, colNameStaff, colNameUrgency, colNameUid, colNameReport, colNameExtraDisplay1, colNameDetails, colNameTags]];
  
  activSpreadsheet.getRange('A1:M1').setValues(dataAry);
}

function getArrayOfAllScheduleSheets()
{
  const sheetsAry = [
  'Farook',
  'Manager',
  'Admin',
  'Svr',
  'Bathu',
  'B5.Ad', 
  'B5.Ic', 
  'B5.Sale', 
  'B5.Dly', 
  'B8.Ad', 
  'B8.Ic', 
  'B8.Sale', 
  'B8.Dly', 
  'B9.Ad', 
  'B9.Ic', 
  'B9.Sale', 
  'B9.Dly', 
  'B11.Ad', 
  'B11.Ic', 
  'B11.Sale', 
  'B11.Dly', 
  'B12.Ad', 
  'B12.Ic', 
  'B12.Sale', 
  'B12.Dly']
  return sheetsAry;
}


function AllSchedulesAutoUpdateStatus(){
  // This function is automatically run by a trigger for 10/30 minutes
  if(!IsTimeBetweenHours(10, 21)) 
    return;
  
  DoAllActionsFromAppSheet();

  const sheetsAry = getArrayOfAllScheduleSheets();
  for (var index = 0; index < sheetsAry.length; index++) {
    AutoUpdateStatusOfSheet(sheetsAry[index]); 
  }
}

function AllSchedulesDeleteEmptyRows(){  
  // This function is automatically run by a trigger daily once
  // This function delete all empty or useless/jobless rows from all
  // schedule sheets
  
  const sheetsAry = getArrayOfAllScheduleSheets();  
  // const sheetsAry = ['Farook', 'Copy of Bathu']

  for (var index = 0; index < sheetsAry.length; index++) {
    
    var sheetName = sheetsAry[index];
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if(sheet == null) return;

    sheet.activate();
    var lastRow = sheet.getMaxRows();

    for(var row = 2; row <= lastRow; row++){            
        if(sheet.getRange(row, colNoWork).getValue())continue;
        if(sheet.getRange(row, colNoRepeat).getValue())continue;
        if(sheet.getRange(row, colNoReport).getValue())continue;
        
        LoggerLog(`Deleting Empty Row: ${row} from Sheet: ${sheetName}`);
        sheet.deleteRow(row);
        lastRow = sheet.getMaxRows();
        row--; // row should be decremented, after delete!
    }

    sheet.insertRowsAfter(lastRow, 5);
    var NewlastRow = sheet.getMaxRows();
    var totalCols = sheet.getMaxColumns();
    sheet.getRange(lastRow, 1, NewlastRow - lastRow + 1, totalCols).setBackground("white");    
    
    DoPageFormatting(sheet);
  }

  // AllSchedulesPageFormatting();
}


function AllSchedulesPageFormatting(){
  const sheetsAry = getArrayOfAllScheduleSheets();  
  
  for (var index = 0; index < sheetsAry.length; index++) {    
    var sheetName = sheetsAry[index];
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if(sheet == null) return;    

    // No worries!
    // It will not hide rows, if user is viewing in Show-All mode
    // setHideDone();
    // refresh2(0, WorkHideDoneWorksNow);  
    // HideColumnsAndSetWidths(sheet);
    // sheet.activate();

    LoggerLog(`DoPageFormatting of Sheet: ${sheetName}`);
    // OnlySetColWidth(sheet);    
    // ;OnlyHideColumns(sheet); 

    DoPageFormatting(sheet);
  }
}


function IsTimeBetweenHours(fromHour, toHour){
  var d = new Date();
  var x = d.getHours();
  return (x >= fromHour) && (x <= toHour);
}

function test333(){
  AutoUpdateStatusOfSheet('Farook');
}

function DoAllActionsFromAppSheet(){  
  LoggerLog("Doing DoAllActionsFromAppSheet():");
  var sheetName = "WorksFromAppSheet";
  var editSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(editSheet == null) return;
  
  const colNoAction = 14;
  const colNoOnSheet = 15;
  
  // This represents ALL the data
  var range = editSheet.getDataRange();  
  var values = range.getValues();

  for (var i = 1; i < values.length; i++) {
  
    var action = values[i][colNoAction-1];
    var onSheetName = values[i][colNoOnSheet-1];
    var jobName = values[i][colNoWork-1];
    var JobDetails = values[i][colNoDetails-1];
    // var status = values[i][colNoStatus-1];
    // var repeat = values[i][colNoRepeat-1];
    // var tags = values[i][colNoTags-1];

    if (jobName == '') continue;
    if (action == '') continue;
    if (onSheetName == '') continue;

    if (action == "AddNewJob") {      
      var onSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(onSheetName);
      CreateNewWorkWithJobName(onSheet, -1, jobName, JobDetails); 
      var txt = `DoAllActionsFromAppSheet(): Sheet: ${sheetName}, AddNewJob Row:${i+1}, Job:${jobName}`;
      LoggerLog(txt);
      editSheet.getRange(i+1, colNoAction).setValue("NewJobAdded");
      continue;
    }
  }
}

function AutoUpdateStatusOfSheet(sheetName){

  // var sheetName = "Farook";
  var editSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(editSheet == null) return;
  
  // var sheetOk = IsItScheduleSheetCheck1(editSheet.getName());  
  // LoggerLog(`AutoUpdateStatusOfSheet():Sheet:${sheetName}, sheetOk: ${sheetOk}`);
  // if (!sheetOk) return;
  
  var sheetOk = IsItScheduleSheetCheck1And2(sheetName);  
  LoggerLog(`AutoUpdateStatusOfSheet():Sheet:${sheetName}, sheetOk: ${sheetOk}`);
  if (!sheetOk) return;



  // This represents ALL the data
  var range = editSheet.getDataRange();  
  var values = range.getValues();

  // This logs the spreadsheet in CSV format with a trailing comma
  for (var i = 1; i < values.length; i++) {
  
    var status = values[i][colNoStatus-1];

    if (status != "" && status != "Disabled" && status != "Over") {
      var editRowNo = i+1;
      var txt = `Found Status in: Row:${editRowNo}, Column:${colNoStatus}`;
      LoggerLog(txt);
      refresh2(editRowNo, WorkUpdate, editSheet);
    }
  }
  
  refresh2(0, WorkHideDoneWorksNow, editSheet);
}


function doGet(e) {

  //https://script.google.com/macros/s/AKfycbwAqlqg19L5GqjWfWPirke64t2-DhUeEyDfnNQZfj3li7M_TvJ4JsXm_W3XQ-1sztjS/exec?email=yes&subject=drive-detected&body=PC3_drive_detected_now
  
  var Uid = e.parameter['uid'];
  if(Uid != null && Uid.length > 0){
    var Table = e.parameter['tablename'];
    UpdateGsheet(Uid, Table);
    return;

    var output = "<html><body style='font-family:tahoma;'><br>" + 
               "Email Sent: " + "<br><br>Subject: " +  "hellosubject" + "<br><br>" + 
               "Body: " + "hellobody" + "</body></html>";
    return HtmlService.createHtmlOutput(output);   
  }
}

function UpdateGsheet(Uid, Table) { 
  var Ss = SpreadsheetApp.openById('1jQ-6typqFmEAk6wtlJKdxlw-yWD7cZo-OuAUsqSrzKI');
  var Actsheet = Ss.getSheetByName(Table); // 'Farook'
  if(Actsheet == null){
    LoggerLog(`doGet():UpdateGsheet(): Table: ${Table}`);  
    return;
  }

  var row = findCell(Uid);
  LoggerLog(`doGet():UpdateGsheet(): row: ${row}`);  
  // return;
  
  var temp = Actsheet.getRange(3, 2, 1,1).getValue();
  Actsheet.getRange(3, 2, 1,1).setValue("Uid:" + Uid + "" + ", row:" + row + ": " + temp );
  LoggerLog(`doGet(): temp: ${temp}`);
}



function randomIntFromInterval(min, max) { 
  // min and max included 
  return Math.floor(Math.random() * (max - min + 1) + min)
}

function showAlert(message){      
   SpreadsheetApp.getUi().alert(message);  
}

function showToast(message){      
  SpreadsheetApp.getActiveSpreadsheet().toast(message);  
}

function onSelectionChange(e){
  if(IsBulkUpdateDoingNow)return;  
  if(GetIsPausedForBulkUpdate()) return; // return during user changes many 'done','not-done'
  
  // var sheetName = SpreadsheetApp.getActiveSheet().getName();
  // var sheetCorrect = IsItScheduleSheetCheck2(sheetName);  
  // if (sheetCorrect == false) return;
  
  var sheetName = SpreadsheetApp.getActiveSheet().getName();
  var correctSheet = IsItScheduleSheetCheck1And2(sheetName);
  LoggerLog(`onSelectionChange():sheetCorrect: ${correctSheet}`);
  if (correctSheet == false) return;


  // execute only once in 10 opportunities, to reduce this call for every user selection change..
  const rndInt = randomIntFromInterval(1, 10)
  if(rndInt < 3){  
    // refresh2(0, WorkHideDoneWorks);    
    // LET'S CHECK THIS FUNCTION'S PERFORMANCE
    
    var bShowAll = getShowAll();
    LoggerLog("getShowAll(): " + bShowAll);  
    if (bShowAll == false) OnlyHideDoneRows();
  }
  else{        
    LoggerLog(`onSelectionChange(): Hide rows in random selection, Not now`);
    ShowJobInExpandedView2();  
  }
}


function PauseForBulkUpdate(Actsheet){
  if (typeof Actsheet === 'undefined') { 
    Actsheet = SpreadsheetApp.getActiveSheet();
  }

  // var Actsheet = SpreadsheetApp.getActiveSheet();
  var rng = Actsheet.getRange(RangeSheetStatus);  
  // rng.setValue("PausedForBulkUpdate");
  rng.setNote("PausedForBulkUpdate"); 

  var sheetStatus = rng.getValue();
  rng.setBackground('#8a0032');
  LoggerLog("sheetStatus: " + sheetStatus + ", Actsheet.Name: " + Actsheet.getName());  
}

function GetIsPausedForBulkUpdate(Actsheet){
  if (typeof Actsheet === 'undefined') { 
    Actsheet = SpreadsheetApp.getActiveSheet();
  }
  // LoggerLog("Actsheet.Name: " + Actsheet.getName());  return;

  // var status = Actsheet.getRange(RangeSheetStatus).getValue();
  var status = Actsheet.getRange(RangeSheetStatus).getNote();
  var paused = status.includes("PausedForBulkUpdate");  
  LoggerLog("PausedForBulkUpdate: " + paused);  
  return paused;
}

function ResumeForBulkUpdate(Actsheet){
  if (typeof Actsheet === 'undefined') { 
    Actsheet = SpreadsheetApp.getActiveSheet();
  }

  // var Actsheet = SpreadsheetApp.getActiveSheet();
  var rng = Actsheet.getRange(RangeSheetStatus);
  // rng.setValue("BulkUpdateOver");
  rng.setNote("");
  var sheetStatus = rng.getValue();
  rng.setBackground('#6aa84f');

  LoggerLog("sheetStatus: " + sheetStatus + ", Actsheet.Name: " + Actsheet.getName());  
}

function test1b(){  
    refresh2(2, WorkUpdate);
}

// This function should be called only after user clicks 'Pause for Bulk Update'
// Then, he updates required works' status
// Then, he clicks 'Do Bulk update now'

function DoBulkUpdate(){
  if(GetIsPausedForBulkUpdate() == false) return; // user would have paused for Bulk update
  if(IsBulkUpdateDoingNow)return;
  
  IsBulkUpdateDoingNow = true;  
  
  var Actsheet = SpreadsheetApp.getActiveSheet();
  var totalRows = Actsheet.getMaxRows();
  Actsheet.setRowHeightsForced(2, totalRows - 1, 1);

  FormatFirstRow(Actsheet, true);
  PauseForBulkUpdate();  // pause other execution for bulk update work.
  
  // var sheetOk = IsItScheduleSheetCheck1(Actsheet.getName());  
  var sheetOk = IsItScheduleSheetCheck1And2(Actsheet.getName());

  LoggerLog("sheetOk: " + sheetOk);
  PrintTimeNow('DoBulkUpdate-IsItScheduleSheetCheck1');    
  if (!sheetOk) return;

  for(let row = StartRow; row <= totalRows; row++){    
    // var newStatus = Actsheet.getRange(row, colNoStatus).isBlank();

    var newStatus = Actsheet.getRange(row, colNoStatus).getValue().toString();
    if(newStatus == "" || newStatus == "Disabled" || newStatus == "Over") {
      LoggerLog("Row:" + row + " is blank/Disabled/Over");
      continue;
    }

    LoggerLog("Calling refresh2() for Row: " + row);
    refresh2(row, WorkUpdate);
  }
  
  IsBulkUpdateDoingNow = false;  
  ResumeForBulkUpdate(); // resume other works
  HideDoneWorks();  
  FormatFirstRow(Actsheet, false);
  Actsheet.setRowHeightsForced(2, totalRows - 1, 30);
  return;
}

function onEdit(e) 
{ 
  if(IsBulkUpdateDoingNow){
    LoggerLog("IsBulkUpdateDoingNow=true, so onEdit exits");
    return;
  }
  if(GetIsPausedForBulkUpdate()) return; // return during user changes many 'done','not-done'

  // return;
  // showAlert('Farook Test message');  
  // e.source.toast('Farook Test message');

  PrintTimeNow('onEdit First Line');
  Logger.log("onEdit Started");
  // This won't work, it returns empty string
  // var email = Session.getActiveUser().getEmail();
  // Logger.log("From onEdit Simple trigger, user email: " + email);  
  
  var activeSheet = e.source.getActiveSheet();
  PrintTimeNow('onEdit-getActiveSheet');

  var activeSheetName = activeSheet.getName();  
  Logger.log("activeSheetName: " + activeSheetName);
  PrintTimeNow('onEdit-activeSheet.getName');  

  // var sheetOk = IsItScheduleSheetCheck1(activeSheetName);  
  // LoggerLog("sheetOk: " + sheetOk);
  // PrintTimeNow('onEdit-IsItScheduleSheetCheck1');    
  // if (!sheetOk) return;

  
  var correctSheet = IsItScheduleSheetCheck1And2(activeSheetName);
  LoggerLog(`onEdit():sheetCorrect: ${correctSheet}`);
  if (correctSheet == false) return;



  var correct = CheckColumnsAreCorrect();
  if (correct != true) 
  {
    e.source.toast("Can't execute.\nColumn Names are wrong", "Danger");
    return;
  }

  //e.source.toast('Refreshing...', '', 1);
  var range = e.range;
  var currentRow = range.getRow();
  var currentCol = range.getColumn();
  var selectedRows = range.getNumRows(); // Returns the number of rows in this range.
  var selectedCols = range.getNumColumns(); // Returns the number of rows in this range.
    
  var cellValue = range.getValue();
  Logger.log(`onEdit(e):AffectedCellValue= ${cellValue}`);

  PrintTimeNow('onEdit-Before move to staff page');  

  if(selectedRows === 1 && selectedCols === 1 && currentCol == colNoReport) 
  {
    Logger.log(`onEdit(e):currentCol= ${currentCol}`);
    // OnlyHideDoneRows();
  }

  //If Staff name is modified then move this row to that staff's page
  //-----------------------------------------------------------------------------
  if(selectedRows === 1 && selectedCols === 1 && currentCol == colNoStaff) 
   {
    var pv = range.getValue();
    if (pv == "") return;

    // //Delete entire row
    // //------------------------------------------------------------------
    // if (pv == "DeleteRow") {    
    //   // var SourceSheetName1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();  
    //   // var valueInColJob =   SpreadsheetApp.getActive().getSheetByName(SourceSheetName1).getRange(row,colNoWork).getValue();    
    //   // valueInColJob = valueInColJob.substring(0,10);
    //   var msg = ui.alert("Sure to delete?", ui.ButtonSet.YES_NO); 
    //     if (msg == ui.Button.YES) {          
    //         var range = SpreadsheetApp.getActive().getSheetByName(activeSheetName).getRange(currentRow,colNoStaff);
    //         range.clearContent();  //I want to delete though row is going to be deleted. If i press ctrl+z then it should not be there. so...
    //         activeSheet.deleteRow(currentRow); 
    //         return;
    //     } else {
    //         var range = SpreadsheetApp.getActive().getSheetByName(activeSheetName).getRange(currentRow,colNoStaff);
    //         range.clearContent();
    //         return;
    //     }  
    // }

    MoveRowToAnotherSheet(currentRow);
    return;
  }

  if(selectedRows === 1 && selectedCols === 1 && currentCol == colNoStatus)  
  {  
    // continue the execution below, after the if-block;
  } 
  else 
  {
    if(selectedRows === 1 && selectedCols === 1 && currentCol == colNoWork)    
    {
      Logger.log("onEdit adding date");
      var jobName = range.getValue();
      Logger.log(`onEdit(e):jobName= ${jobName}`);
      if (!jobName) return;     
      
      CreateNewWorkWithJobName(activeSheet, currentRow, jobName); 
      return;
    }

    Logger.log("onEdit exited,selectedRows/cols !== 1 or colNoStatus not changed.");
    return;
  }
  
  var oldStr = e.oldValue;
  var newStr = e.value;    
  oldStr = (oldStr == null)? "" : oldStr.toString();
  newStr = (newStr == null)? "" : newStr.toString();
  Logger.log("oldStr: " + oldStr + ",  newStr: " + newStr);   
  PrintTimeNow('onEdit20-Before calling refresh2()');

  if (newStr.includes("Disabled") || newStr.includes("Alloted")) // .contains()
  {
    refresh2(currentRow, WorkUpdate);      
    Logger.log("Good, changed to Alloted or Disabled");   
  }
  else if (newStr.includes("Done") || newStr.includes("Worked") || newStr.includes("Not Done"))
  {
    // showToast("Use report column to add any info!");
    Logger.log("Good, will change to Done/NotDone");   
    refresh2(currentRow, WorkUpdate);
  }
  else if (newStr.includes("Postpone"))
  {
    refresh2(currentRow, WorkUpdate, activeSheet);
    Logger.log("Good, Just postponed");   
  }
  else if (newStr.includes("----"))
  { 
    Logger.log("Seperator pressed");   
    return;
  }
  else if (newStr.includes("Delete"))
  { 
    var workName3 = activeSheet.getRange(currentRow, colNoWork).getValue();
    var rngFrom = activeSheet.getRange(currentRow, 1, 1, colNoReport);
    var copiedRowValues = rngFrom.getValues();
    var sheetDelWorks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetWorksDeleted);
    sheetDelWorks.insertRowBefore(1);
    var pasteRng = sheetDelWorks.getRange(1, 1, 1, colNoReport);
    pasteRng.setValues(copiedRowValues);    
    var pasteRng2 = sheetDelWorks.getRange(1, colNoReport +1);
    pasteRng2.setValue(activeSheet.getName());
    Logger.log("Deleted work: " + workName3);   
    activeSheet.deleteRow(currentRow);
    return;
  }
  else
  {
    refresh2(0, WorkHideDoneWorks);      
    Logger.log("Not changed to Done");    
  }


  /* Copy Job time to Postponed time when add new job
   var sheet = SpreadsheetApp.getActiveSheet();
  if (currentCol == colNoProperTime)
  {    
    Logger.log("Bathu1 : " + currentCol)
    Logger.log("Bathu2 : " + sheet.getRange(currentRow, colNoProperTime).getValue)
    var myJobTime = sheet.getRange(currentRow, colNoProperTime).getValue
    sheet.getRange(currentRow, colNoDue).setValue(myJobTime);   
  }
  */

  
  Logger.log("onEdit finished");        
  //e.source.toast('Refresh Done...', '', 5);
}


function CreateNewWorkWithJobName(activeSheet, currentRow, jobName, JobDetails = ''){

  if(activeSheet == null) return;
  var actSheetName = activeSheet.getName();

  // var sheetOk = IsItScheduleSheetCheck1(actSheetName);   
  var sheetOk = IsItScheduleSheetCheck1And2(actSheetName);

  LoggerLog(`CreateNewWorkWithJobName():Sheet:${actSheetName}, sheetOk: ${sheetOk}`);
  if (!sheetOk) return;

  // if new job is added without user's direct interation with the sheet
  // i.e, when adding a new work from another sheep or Appsheet App
  if(currentRow == -1){
    currentRow = GetFirstEmptyRow(activeSheet);
    LoggerLog(`New job will be created on Row: ${currentRow}, in Sheet:${actSheetName}`);
    
    var rngJob = activeSheet.getRange(currentRow, colNoWork);
    var workName = rngJob.getValue();
    if (!workName){
      if(jobName) rngJob.setValue(jobName);
    }
  }

  // Proper Date
  var dt = new Date();
  var rngDateProper = activeSheet.getRange(currentRow, colNoProperTime);
  var dp = rngDateProper.getValue();
  // if (dp.length != 0) return;      
  if (dp.length == 0) rngDateProper.setValue(dt);

  // Due Date
  rngDateDue =  activeSheet.getRange(currentRow, colNoDue);
  var dd = rngDateDue.getValue();
  // if (dd.length != 0) return;      
  if (dd.length == 0) rngDateDue.setValue(dt);

  // Report
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yy")
  rngReport = activeSheet.getRange(currentRow, colNoReport);            
  rngReport.setValue(date + ' Created' + "\n" + rngReport.getValue());
  
  // Uid
  var rngUid = activeSheet.getRange(currentRow, colNoUid);
  if(!rngUid.getValue()){    
    var newUid = GetUniqueId();
    rngUid.setValue(newUid);
  }

  // Repeat
  var rngRepeat = activeSheet.getRange(currentRow, colNoRepeat);            
  if(!rngRepeat.getValue()){
    rngRepeat.setValue("Once")  
    rngRepeat.activate(); 
  }

  // JobDetails
  var rngDetails = activeSheet.getRange(currentRow, colNoDetails);           
  if(JobDetails){
    if(rngDetails.getValue()) JobDetails = JobDetails + "\n" + rngDetails.getValue();
    rngDetails.setValue(JobDetails)  
  }

  // Remove Tags  
  var rngTags = activeSheet.getRange(currentRow, colNoTags);           
  if(rngTags.getValue() == WorkNew){
    rngTags.setValue(""); 
  }

}

function IsItScheduleSheetCheck2(sheetName){
  LoggerLog("IsItScheduleSheetCheck2: Active Sheet Name: " + sheetName);

  if (sheetName == "Base") return false;
  if (sheetName == "Base1") return false;
  if (sheetName == "DayJob") return false;
  if (sheetName == "Sales") return false;
  if (sheetName == "Kids") return false;
  if (sheetName == "WorksDone") return false;
  if (sheetName == "Done") return false;
  if (sheetName == "RpDua") return false;
  if (sheetName == "Quran") return false;
  if (sheetName == "Time") return false;
  if (sheetName == "Bank") return false;
  if (sheetName == "VDoRep") return false;
  if (sheetName == "Db") return false;
  if (sheetName == "Info") return false;
  if (sheetName == "Idea") return false;
  if (sheetName == "ClrStk") return false;
  if (sheetName == "Temp") return false;
  if (sheetName == "logs") return false;
  if (sheetName == "logs2") return false;
  if (sheetName == "logs3") return false;
  if (sheetName == "logs4") return false;
  if (sheetName == "Msg") return false;
  if (sheetName == "Feedback") return false;
  if (sheetName == "WorksFromAppSheet") return false;
  if (sheetName == "AppSheetInstructions") return false;  

  return true;
}


function Check_SheduleSheet(){
  // JUST FOR CHECKING PURPOSE
  var activeSheetName = SpreadsheetApp.getActive().getActiveSheet().getName();
  var result = IsItScheduleSheetCheck1And2(activeSheetName);
  LoggerLog(`Result: ${result}`);
}

function navigateToLastPosition() {
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();
  cursor.setPosition(doc.getLastPosition());
}

function refresh2(rowEdited, whatWorkToDo, sheet)
{
  // HELP: THIS FUNCTION DOES THE FOLLOWING
  // 1) IF USER CHANGES WORK'S STATUS, IT CHANGES DUETIME
  // 2) REFRESHES THE SHOWN/HIDDEN WORKS EVERY 10 MINUTES

  NoteDownThisTimeAsOldTime();

  if (typeof sheet === 'undefined') { 
    var sheet = SpreadsheetApp.getActiveSheet();
  }
  // var sheet = SpreadsheetApp.getActiveSheet();
  
  var sheetName = sheet.getName();  
  // var sheetOk = IsItScheduleSheetCheck1(sheetName);   
  var sheetOk = IsItScheduleSheetCheck1And2(sheetName);

  LoggerLog("refresh2: sheetOk: " + sheetOk);
  if (!sheetOk) return;

  // var sheetCorrect = IsItScheduleSheetCheck2(sheetName);  
  // PrintTimeNow('stage1');
  // if (sheetCorrect == false) return;

  


  // var pause = PropertiesService.getScriptProperties().getProperty('KeyPauseExecution');  
  // Logger.log("Not used, Pause: " + pause);
  // if (pause) return;

  // If bShowAll=true, unhide all rows
  var bShowAll = getShowAll();
  LoggerLog("getShowAll(): " + bShowAll);  
  
  var timeNow = SpreadsheetApp.getActive().getRangeByName("Db!A3").getValue();    
  LoggerLog("Refresh2, rowEdited: " + rowEdited);
  var changeDueTime = false;  
    
  PrintTimeNow('stage2');
  // var oldtime = new Date();

  var isAbove10Mins = IfExecutedTimeAboveMinutes(10);      

  // HIDE/UNHIDE LOGIC SECTION STARTS
  // Set to true to update only the edited row, not all rows.  
  var ToUpdateSingleRow = false; // decides should update all rows? or only the edited row?
  ToUpdateSingleRow = !isAbove10Mins;

  if(IsBulkUpdateDoingNow) {
    isAbove10Mins = false;     // it is not neccessary
    ToUpdateSingleRow = true;  // this refresh function should be executed fully for each row
    whatWorkToDo = WorkUpdate; // It is WorkUpdate, not WorkHideDoneWorks!
  }

  if (whatWorkToDo == WorkHideDoneWorks)   
  {     
    // Do HideDoneWorks only every 10 minutes
    changeDueTime = false;
    
    if (isAbove10Mins == false) 
    { 
      // If less than 10 mins
      // Show(Display) job in expanded view
      ShowJobInExpandedView2();      
      return;
    }
  }
  else if (whatWorkToDo == WorkHideDoneWorksNow) 
  {
    // This is executed when menu "Hide done works" is clicked.
    // Do HideDoneWorks immediately, without checking last executed time for HideDoneWorks
    // Don't update the edited row
    changeDueTime = false;
    ToUpdateSingleRow = false;
  }
  else  
  {    
    // this block works when (whatWorkToDo == WorkUpdate) 
    // Change DueTime, so update the edited row
    changeDueTime = true;
  }
  
  PrintTimeNow('stage3');  
  if (sheetName == "Base") return;

  var totalCols = sheet.getMaxColumns();
  var totalRows = sheet.getMaxRows();

  var rng = sheet.getRange(StartRow, 1, totalRows - StartRow + 1, totalCols);
  var rowAry = rng.getValues();  
  var totalRowsInAry = rowAry.length - 1;
  var postponed = false;
  // var rngDate = sheet.getRange(StartRow, 1);
  PrintTimeNow('stage4');  

  // FIND NEW DUETIME
  // CHANGE-DUETIME LOGIC SECTION STARTS
  if (changeDueTime) // This is true only when 'whatWorkToDo == WorkUpdate'
  {    
    var rowEditedForAry = rowEdited - StartRow;
    var properTimeStr = rowAry[rowEditedForAry][colNoProperTime - 1];
    var properTime = new Date(properTimeStr);
    var dueTimeStr = rowAry[rowEditedForAry][colNoDue - 1];
    var dueTime = new Date(dueTimeStr);    
    var frequency = rowAry[rowEditedForAry][colNoRepeat - 1];    
    var jobName = rowAry[rowEditedForAry][colNoWork - 1];        
    Logger.log(`onEdit(e):jobName= ${jobName}`);

    var rangeProperTime = sheet.getRange(rowEdited, colNoProperTime);   
    var rangeStatus = sheet.getRange(rowEdited, colNoStatus);  
    var rangeRepeat = sheet.getRange(rowEdited, colNoRepeat);  
    var rangeDue = sheet.getRange(rowEdited, colNoDue);   
    var newStatus = rangeStatus.getValue();
    newStatus = (newStatus == null)? "" : newStatus.toString();
    var postponed = newStatus.includes("Postpone"); // bool
    
    // ADD YES/NO IN REPORT COLUMN FOR DONE/NOT-DONE STATUS CHANGES
    if (newStatus == "Done" || newStatus == "Worked"){
      var dt = new Date();
      var nt = dt.getDate()
      var colNoOfTodayWork = colNoReport + nt;
      // var rangeOfTodayWork = sheet.getRange(rowEdited, colNoOfTodayWork);  //Now I dont put tick on todays date. 
      // rangeOfTodayWork.setValue("Yes: " + dt.toLocaleString());  
      var rangeOfReport = sheet.getRange(rowEdited, colNoReport);   
      
      // var ValueInRangeReport = text.RangeofReport.trim();
      //  var minutesStr = text.replace(/\D/g,'').trim();

      // var rr = rangeOfReport.getValue();
      // Logger.log ("RangeofReport " + rr );
      // var txt1 = convertDate(dt)  + " Yes" + "\n" + rangeOfReport.getValue() ;
      
      // I add empty line at first. So each row will looks neat.
      var addTxt = (newStatus == "Done")? "Yes" : "Worked";
      // var txt1 = convertDateShort(dt)  + " Yes" + "\n" + rangeOfReport.getValue().trim() ;
      var txt1 = convertDateShort(dt)  + " " + addTxt + "\n" + rangeOfReport.getValue().trim() ;
      
      rangeOfReport.setValue(txt1); //It will have all done days report
    }

    if (newStatus == "Not Done"){
      var dt = new Date();
      var nt = dt.getDate()
      var colNoOfTodayWork = colNoReport + nt;
      // var rangeOfTodayWork = sheet.getRange(rowEdited, colNoOfTodayWork);  //Now I dont put tick on todays date. 
      // rangeOfTodayWork.setValue("No: " + dt.toLocaleString());  
      var rangeOfReport = sheet.getRange(rowEdited, colNoReport);   
      var rr=rangeOfReport.getValue();
      Logger.log ("RangeofReport " + rr );
      // var txt1 = convertDate(dt)  + " No" + "\n" + rangeOfReport.getValue() ;
      
      // I add empty line at first. So each row will looks neat.
      var txt1 = convertDateShort(dt)  + " No" + "\n" + rangeOfReport.getValue().trim() ;
      rangeOfReport.setValue(txt1); //It will have all done days report
    }


    if (frequency == "Once" && postponed == false)
    {
      // rangeStatus.setValue("Done"); // Previously it was left un-altered.
      rangeRepeat.setValue("OnceDone");      
      rangeStatus.setValue("Over");    // Now, change to "Over"
    }
    else
    {
        rangeStatus.setValue("");
    }
    
    var today = new Date();        
    var todayLastSecond = today.setHours(23,59,59,0);     
    var newDueTime = properTime;    

    // PUT NEW DUETIME FOR CURRENTLY EDITED ROW ONLY!!
    switch (frequency) {

      case "Once":      
        if (postponed)
        {
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }
        newDueTime = new Date();     
        break;
        
      case "Hourly":
      case "A-Hourly":      
        if (postponed)
        {
          //newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          newDueTime = getDaysToPostponeFromText(newStatus, new Date(timeNow));
          break;
        }

        var today = new Date(); 
        do 
        {
          newDueTime = addMinutes(newDueTime, 60);  
        } 
        while (newDueTime < timeNow);
        break;

      case "Daily":      
      case "B-Daily":      
        if (postponed)
        {
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = newDueTime.addDays(1);  
        } 
        while (newDueTime < todayLastSecond);
        break;

      case "Weekly":
      case "C-Weekly":
        if (postponed)
        {
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = newDueTime.addDays(7);  
        } 
        while (newDueTime < todayLastSecond);
        break;

      case "Biweekly":
      case "D-BiWeekly":
        if (postponed)
        {
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = newDueTime.addDays(14);  
        } 
        while (newDueTime < todayLastSecond);
        break;

      case "Monthly":
      case "E-Monthly":     

        if (postponed){
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = addMonths(newDueTime, 1);  
        } 
        while (newDueTime < todayLastSecond);        
        break;


      case "Bimonthly":
      case "F-BiMonthly":     

        if (postponed){
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = addMonths(newDueTime, 2);  
        } 
        while (newDueTime < todayLastSecond);        
        break;

      case "Quaterly":
      case "G-Quaterly":     

        if (postponed){
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = addMonths(newDueTime, 3);  
        } 
        while (newDueTime < todayLastSecond);        
        break;


      case "4Monthly":
      case "H-4Monthly":     

        if (postponed){
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = addMonths(newDueTime, 4);  
        } 
        while (newDueTime < todayLastSecond);        
        break;

      case "Halfyearly":
      case "I-HalfYearly":     

        if (postponed){
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = addMonths(newDueTime, 6);  
        } 
        while (newDueTime < todayLastSecond);        
        break;

      case "Yearly":
      case "J-Yearly":        

        if (postponed)
        {
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = addYears(newDueTime, 1);  
        } 
        while (newDueTime < todayLastSecond);        
        break;
        

      case "Biyearly":
      case "K-BiYearly":        

        if (postponed)
        {
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = addYears(newDueTime, 2);  
        } 
        while (newDueTime < todayLastSecond);        
        break;

      case "3Yearly":
      case "L-3Yearly":        
        if (postponed)
        {
          newDueTime = getDaysToPostponeFromText(newStatus, new Date());
          break;
        }

        do 
        {
          newDueTime = addYears(newDueTime, 3);  
        } 
        while (newDueTime < todayLastSecond);        
        break;
        
      
      default:
      Logger.log("Frquency is wrong: " + frequency);
      SpreadsheetApp.getActiveSpreadsheet().toast("Frquency is wrong"  + frequency, "", 1);      
      return;
    }
    
    if (postponed)
    {      
      rangeDue.setValue(newDueTime);        
      rowAry[rowEditedForAry][colNoDue - 1] = newDueTime;
      Logger.log("In Postponed If block, newDueTime: " + newDueTime);
    }
    else // if work Done
    {
      rangeDue.setValue(newDueTime);    
      rowAry[rowEditedForAry][colNoDue - 1] = newDueTime;
      rangeProperTime.setValue(newDueTime);
      rowAry[rowEditedForAry][colNoProperTime - 1] = newDueTime;      
    }
    
  // SpreadsheetApp.getActiveSpreadsheet().toast("Refresh, changeDueTime: " + changeDueTime, "", 6);
  }  
  // CHANGE-DUETIME LOGIC SECTION ENDS

  LoggerLog("ToUpdateSingleRow =  " + ToUpdateSingleRow + ",  whatWorkToDo=" + whatWorkToDo);
  PrintTimeNow('stage5');
  
  // HIDE/UNHIDE-ROWS LOGIC SECTION STARTS
  // single row (currently edited row) or all rows (ToUpdateSingleRow?)
  for (let row = 0; row < totalRowsInAry+1; row++) //totalRowsInAry
  {

    if (ToUpdateSingleRow){      
      row = rowEditedForAry;
    }

    var status = rowAry[row][colNoStatus - 1];
    var workName = rowAry[row][colNoWork - 1];    
    
    var properTimeStr = rowAry[row][colNoProperTime - 1];
    var properTime = new Date(properTimeStr);

    var dueTimeStr = rowAry[row][colNoDue - 1];
    var dueTime = new Date(dueTimeStr);

    var frequency = rowAry[row][colNoRepeat - 1];
    var actualRow = row + StartRow;    

    // PrintTimeNow(oldtime, 'stage5x-' + row.toString());
    // var oldtime = new Date();

    //var d = new Date();
    //var timeNow = d.getTime();     
    //var timeNow = d;     
    //var dstr = d.toLocaleString();
    //Logger.log("toLocaleDateString : "+  dstr);
    //timeNow = new Date(dstr);

    // LoggerLog(  "Row: " + actualRow + ", Wname: " + workName.substring(0, 10) + ", Freqn: " + frequency + ", Status: " + status +  ", dueTimeStr: " + dueTimeStr + ", dueTime: " + dueTime + ", timeNow: " + timeNow.toString());

    var rangex = sheet.getRange(actualRow,1, 1, totalCols);
    // var rangeUrgency = sheet.getRange(actualRow,colNoUrgency, 1, 1); // should be removed
    var rangeTags = sheet.getRange(actualRow,colNoTags, 1, 1);

    if ((frequency == "Once" && (status == "Done" || status == "Over")) || 
         dueTime > timeNow ||
         status == "Disabled" || status == "Alloted" )
    {
      rangex.setBackground("#e8e8e8");      // #c3d7e3 #f2f2f2 #e8e8e8
      // rangeUrgency.setValue("");  // should be removed
      rangeTags.setValue("");
      

      // SINCE BATHU DOESN'T WANT THE ROW TO BE HIDDEN IMMEDIATELY AFTER
      // USER PUTS 'DONE/UNDONE/POSTPONE/ANYTHING'
      // FAROOK HAS PUT A NEW FUNCTION OnlyHideDoneRows()
      // WHICH HIDES DONE ROWS AT SOMETIME! CHECK CODE FOR THAT.
      // skip hiding the row, if the user is viewing in ShowAll mode
      // if(!bShowAll) sheet.hideRows(actualRow);

      // Logger.log(workName.substring(0, 10) + " hidden now: " + actualRow);

      // var dtNow = new Date();
      // sheet.getRange(actualRow, colNoEditedTime).setValue(dtNow);

    } 
    else
    {
      // Logger.log(workName.substring(0, 10) + " visible now: " + actualRow);
      // var rangeToUnHide = rangex;
      // sheet.unhideRow(rangeToUnHide);   
      // sheet.isRowHiddenByUser(actualRow);

      if(!workName) continue;
      sheet.unhideRow(rangex);
      rangex.setBackground("white");      
      rangeTags.setValue("show");
      // rangeUrgency.setValue("*");  // should be removed
    }

    if (ToUpdateSingleRow){  
      LoggerLog("ToUpdateSingleRow=True, row=" + row);
      break;
    }
  }
  // HIDE/UNHIDE-ROWS LOGIC SECTION ENDS

  PrintTimeNow('stage6');

  // Last ExecutedTime (hidden,unhidden time) is required only when all rows are processed
  if (ToUpdateSingleRow == false)SaveExecutedTime(timeNow);

  // showDoneToast(changeDueTime, postponed);
  PrintTimeNow('stage7');
}

function OnlyHideDoneRows(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // ClearJobReport(sheet);

  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][colNoTags-1] === "") { // Check if the 4th column (column D) is empty
      sheet.hideRows(i + 1); // Rows are 1-based, so we add 1 to the index
    }else{
      sheet.unhideRow(sheet.getRange(i + 1,1));
    }
  }
}


function showDoneToast(changeDueTime, postponed){
  Utilities.sleep(5000);
  Logger.log("Showing Refresh Done Toast...");
  return;

  if(changeDueTime){
    SpreadsheetApp.getActiveSpreadsheet().toast("Refresh Done", "Done", 6);  
  }else{
    SpreadsheetApp.getActiveSpreadsheet().toast("Refresh Done." + "\n" + 
    "\nPostponed: " + postponed + "\nchangeDueTime: " + changeDueTime, "Done", 6);        
  }    
  //SpreadsheetApp.getUi().alert("Alert message");
}


function ShowJobInExpandedView1(){
    // Show(Display) job in expanded view
    // ToUpdateSingleRow = true;

    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName(); 
    // LoggerLog("sheetName: " + sheetName);
    if (sheetName != "Farook" ) return;

    // sheet.setFrozenRows(2);
    if(sheet.getCurrentCell().getColumn() != colNoWork) return;
    sheet.getRange(2, 3).setValue(sheet.getRange(sheet.getCurrentCell().getRow(), colNoWork).getValue());    

}

function ShowJobInExpandedView2(){    

    if(IsExtraDisplayEnabled == false) 
      return;
    
    // Show(Display) job in expanded view
    // ToUpdateSingleRow = true;
    // if (sheetName == "Farook" || sheetName == "Bathu" || sheetName == "Manager")   

    var sheet = SpreadsheetApp.getActiveSheet();
    var rowEdited = sheet.getCurrentCell().getRow();
    var colEdited = sheet.getCurrentCell().getColumn();
    //SpreadsheetApp.getUi().alert("A");
    if(colEdited != colNoWork) return;

    var rngExtraDisplay1 = sheet.getRange(rowEdited, colNoExtraDisplay1);        
    rngExtraDisplay1.setWrap(true);

    rngExtraDisplay1c = sheet.getRange("K:K");          
    rngExtraDisplay1c.breakApart();         
    rngExtraDisplay1c.clearContent();  

    var fd1 = sheet.getCurrentCell().getValue(); 
    rngExtraDisplay1b = sheet.getRange(rowEdited, colNoExtraDisplay1, 40, 1);  
    rngExtraDisplay1b.merge();
    rngExtraDisplay1b.setValue(fd1);     
    rngExtraDisplay1b.setVerticalAlignment("top");    
    rngExtraDisplay1b.setFontSize(12);   
    rngExtraDisplay1b.setFontFamily("Verdana");
    rngExtraDisplay1b.setBackground("#ffffff"); //not works. but yellow red works
    sheet.setColumnWidth(colNoExtraDisplay1, 600); //Info 
    LoggerLog("E3 Test W234: " + fd1);  
}

function onOpen() {  
  
  var Actsheet = GetActiveSheet();
  PutUidForAllJobs(Actsheet);
  DoBulkUpdate();
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SunCo')
    // .addItem('Clear Job Report', 'ClearJobReport')
    .addItem('Create Job Report', 'CreateJobReport')
    // .addItem('Create Pdf', 'generatePdf')
    .addItem('Create Job Report & Pdf', 'CreateJobReportAndPdf')
    .addSeparator()
    .addItem('Create Other Month JR & Pdf', 'CreateJobReportAndPdfForMonth')
    // .addItem('Download PDF', 'openUrl1')
    .addSeparator()
    .addItem('Show All Works', 'showAllWorks')
    .addItem('Hide Done Works', 'HideDoneWorks')
    // .addItem('Sort By Frequency', 'SortByFrequency')
    .addItem('Page Formatting', 'PageFormatting')
    // .addItem('Page Formatting All', 'AllSchedulesDeleteEmptyRows')    
    .addSeparator()
    // .addItem('Pause Execution', 'pauseExecution')
    // .addItem('Resume Execution', 'resumeExecution')
    .addItem('Pause For BulkUpdate', 'PauseForBulkUpdate')    
    .addItem('Do BulkUpdate Now', 'DoBulkUpdate')    
    .addSeparator()
    .addItem('Display Finished Works', 'DisplayFinishedWorks')
    .addItem('Show Admin Sidebar', 'showAdminSidebar')
    .addItem('Check IfAnyWork Missing...', 'CheckIfAnyWorkMissing')
    // .addItem('Put Uid For AllJobs in ColInfo', 'PutUidForAllJobs')
     // .addItem('createMissingColumnsTemp3', 'createMissingColumnsTemp3')    
    .addItem('Test E3', 'E3Test')
    .addToUi();
    
    // .addItem('--------------------','mnuSeperator')    
    // .addItem('--------------------','mnuSeperator')
    
    // Utilities.sleep(10000);
    // showAdminSidebar();

    // PageFormatting();
    // fnLockRange();
    
    HideDoneWorks();
    
}


function createSpreadsheetOpenTrigger() 
{
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onEditAdvanced')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
}


function mnuSeperator() {
  return
}



function fnLockRange() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I:K').activate();
  var protection = spreadsheet.getRange('I:K').protect();
  protection.setDescription('rngLocked');
};


function pauseExecution() {
  PauseExecution = true;
  Logger.log("PauseExecution: " + PauseExecution);
  PropertiesService.getScriptProperties().setProperty('KeyPauseExecution', true);
}

function resumeExecution() {
  PauseExecution = false;
  Logger.log("PauseExecution: " + PauseExecution);
  PropertiesService.getScriptProperties().setProperty('KeyPauseExecution', false);
}

function DelCheckHiddenRow(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];

  // Rows start at 1
  Logger.log(sheet.isRowHiddenByUser(3));
}

function DelCheckDayOfToday(){
  var dt = new Date();
  var nt = dt.getDate()  
  Logger.log("nt: " + nt);
  Logger.log("dt: " + dt);
  var colNoOfTodayWork = colNoStatus + nt;

  Logger.log("colNoOfTodayWork: " + colNoOfTodayWork);
  var sheet = SpreadsheetApp.getActiveSheet();
  var rangeOfTodayWork = sheet.getRange(2, colNoOfTodayWork);   
  Logger.log("rangeOfTodayWork: " + rangeOfTodayWork.getA1Notation());
}

function check1(){  
  
  var sn2 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  Logger.log("Active Sheet Name2: " + sn2);  
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();
  Logger.log("Active Sheet Name: " + sheetName);  
  // if (sheetName != "Base")     return;
  return;

  var today = new Date();   
  Logger.log("check1, today1: " + today);
  Logger.log("check1, today1: " + today.toLocaleString());
  var t = today.getTime();
  var ac = SpreadsheetApp.getActive().getRangeByName("Db!F2").getValue();
  Logger.log("check1, ac: " + ac);
  Logger.log("check1, today1: " + t);



  return;
  var todayx = today.setHours(23,59,59,0);
  var todayx2 = new Date(todayx);
  Logger.log("check1, today2: " + today.toString());
  Logger.log("check1, todayx: " + todayx.toString());
  Logger.log("check1, todayx2: " + todayx2.toString());
}

function displayToast(who) {
  SpreadsheetApp.getActive().toast("Hi there!");  
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(2,1).setValue(who);
}

function NewJob(){
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var shtName = sheet.getName();
  sheet.insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
  var rowPrev = spreadsheet.getActiveRange().getRow();
  
  Logger.log("NewJob, rowPrev: " + rowPrev);
  if(rowPrev == 1) {
    rowPrev = 2
    LoggerLog("NewJob, rowPrev: " + rowPrev + ", changed to 2");
  }

  var rowNew = rowPrev + 1;
  var rng = sheet.getRange(rowNew, colNoSNo);
  var rng2 = sheet.getRange(rowNew, colNoWork);
  var rngDateProper = sheet.getRange(rowNew, colNoProperTime);
  var rngDateDue = sheet.getRange(rowNew, colNoDue);
  var rngRepeat = sheet.getRange(rowNew, colNoRepeat);
  var rngInfo = sheet.getRange(rowNew, colNoUid);
  Logger.log("NewJob, rng2: " + rng2.getA1Notation());

  rng.setValue(1);
  rng2.setValue("Type job & edit other data");  
  var dt = new Date();
  rngDateProper.setValue(dt);
  rngDateDue.setValue(dt);  
  var newUid = GetUniqueId();
  rngInfo.setValue(newUid);
  rngRepeat.setValue("Once")  

  if(shtName == "Base"){    
    var uId = GetUniqueIDForNewJob2(sheet);
    if(uId == -1){
      showToast("Please check if any Id is missing in Base Sheet");
    }
    else{
      rng.setValue(uId);
    }
  }
  rng2.activate();
  
  Logger.log("NewJob, rngRepeat: " + rngRepeat.getA1Notation());
}


function AssignEmp(empName){

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();  // Need to get the sheet, not just the whole work book.
  var name1 = sheet.getSheetName();
  Logger.log("SheetName1: " + name1);  

  var ac = sheet.getActiveCell();  
  var row = ac.getRow();
  sheet.getRange(row, colNoStaff).setValue(empName);


  return;
  for(let row = StartRow; row < 10; row++){
    var empassigned = sheet.getRange(row, colNoStaff).getDisplayValue();
    Logger.log("row: " + row.toString() + ", empassigned: " + empassigned);  
  }
}

function SortByFrequency() {

  var sheet = SpreadsheetApp.getActiveSheet();
  // SortColumnByJobFrequency(sheet);
  SortColumnsBy462(sheet);
  return;


  // Utilities.sleep(5000);
  var row = 5;
  var sheet = SpreadsheetApp.getActiveSheet();
  var totalRows = sheet.getMaxRows();
  var totalCols = sheet.getMaxColumns();

  var r1 = sheet.getRange(1, colNoUid, 1, 1);
  r1.setValue("A");

  for (let row = StartRow; row < totalRows - StartRow; row++)
  {    
    var frequency = sheet.getRange(row , colNoRepeat, 1, 1).getDisplayValue();
    var rngSort = sheet.getRange(row , colNoUid, 1, 1);

    switch (frequency) {
    case "Once":      
      rngSort.setValue("O");
      break;
    case "Hourly":
    case "A-Hourly":      
      rngSort.setValue("C");
      break;
    case "Daily":
    case "B-Daily":      
      rngSort.setValue("D");
      break;
    case "Weekly":
    case "C-Weekly":      
      rngSort.setValue("E");
      break;
    case "Biweekly":
    case "D-Biweekly":      
      rngSort.setValue("F");
      break;
    case "Monthly":
    case "E-Monthly":      
      rngSort.setValue("G");  
      break;
    case "Bimonthly":
    case "F-BiMonthly":      
      rngSort.setValue("H");
      break;
    case "Quaterly":
    case "G-Quaterly":      
      rngSort.setValue("I");
      break;
    case "4Monthly":
    case "H-4Monthly":      
      rngSort.setValue("J");
      break;
    case "Halfyearly":
    case "I-HalfYearly":      
      rngSort.setValue("K");
      break;
    case "Yearly":
    case "J-Yearly":      
      rngSort.setValue("L");
      break;
    case "Biyearly":
    case "K-BiYearly":      
      rngSort.setValue("M");
      break;
    case "3Yearly":
    case "L-3Yearly":      
      rngSort.setValue("N");
      break;      
    case "":      
      rngSort.setValue("");
      break;

    default:
      rngSort.setValue("Z");
    }
  }
  

};
 


function SortColumnByJobFrequency(sheet) {
  sheet.getRange('D:D').activate();
  sheet.sort(4, true); 
};



function SortColumnsBy462(Actsheet) {

  if (typeof Actsheet === 'undefined') { 
    Actsheet = SpreadsheetApp.getActiveSheet();
  }

  // var Actsheet = SpreadsheetApp.getActiveSheet();

  // var sheetOk = IsItScheduleSheetCheck1(Actsheet.getName());  
  var sheetOk = IsItScheduleSheetCheck1And2(Actsheet.getName());

  LoggerLog("sheetOk: " + sheetOk);
  if (!sheetOk) return;

  if(IsExtraDisplayEnabled) Actsheet.getRange("K:K").clearFormat();
  var totalRows = Actsheet.getMaxRows();
  
  SORT_ORDER = [
    {column: 6, ascending: true},
    {column: 2, ascending: true},
    {column: 3, ascending: false},
    {column: 4, ascending: true},
  ];


  colTotal = Actsheet.getMaxColumns();
  var range = Actsheet.getRange(2,1,totalRows-1,colTotal);
  // range.sort(SORT_ORDER);

  Actsheet.sort(6, true);
  // mySleep(2);
  Actsheet.sort(2, true);
  // mySleep(2);
  Actsheet.sort(3, false);
  // mySleep(2);
  Actsheet.sort(4, true);
 
  /**/

};

function TestSort1(){
  var Actsheet = SpreadsheetApp.getActiveSheet();
  var totalRows = Actsheet.getMaxRows();
  colTotal = Actsheet.getMaxColumns();

  SORT_ORDER = [
    {column: 3, ascending: true},
  ];

  var range = Actsheet.getRange(2,1,totalRows-1,colTotal);
  range.sort(SORT_ORDER);
}



function nouse(){
  /*
  const colNoSNo = 1;
  const colNoWork = 2;
  const colNoProperTime = 3;
  const colNoDue = 4;
  const colNoRepeat = 5;
  const colNoStaff = 6;
  const colNoUrgency = 7;
  const colNoStatus = 8;
  const colNoUid = 9;
  const colNoReport = 10;
  */

  /*
  const colNoSNo = 1;
  const colNoWork = 2;
  const colNoStatus = 3;

  const colNoProperTime = 4;
  const colNoDue = 5;
  const colNoRepeat = 6;
  const colNoStaff = 7;
  const colNoUrgency = 8;

  const colNoUid = 9;
  const colNoReport = 10;
  const StartRow = 2;
 */
}

function CheckColumnsAreCorrect(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var totalCols = sheet.getMaxColumns();
  var rng = sheet.getRange(1, 1, 1, totalCols);
  var rowAry = rng.getValues();  
  // SNo,	Work Name,	Proper Time, Time,	Repeat,	Staff,	Urgency,	Status


  // the below cell is used for 'Job-show' 'Job-hide' status
  // if (rowAry[0][colNoSNo - 1] != "SNo") return false;          // 1 - A - Sno(Id)
  if (rowAry[0][colNoWork - 1] != "Job") return false;            // 2 - B - JobName
  if (rowAry[0][colNoStatus - 1] != "Status") return false;       // 3 - C - Status
  if (rowAry[0][colNoRepeat - 1] != "Repeat") return false;       // 4 - D - Repeat

  if (rowAry[0][colNoProperTime - 1] != "Job time") return false; // 5 - E - Job Time
  if (rowAry[0][colNoDue - 1] != "Postponed") return false;       // 6 - F - Postponed Time
  if (rowAry[0][colNoStaff - 1] != "Staff") return false;         // 7 - G - Staff
  if (rowAry[0][colNoUrgency - 1] != "Urgency") return false;     // 8 - H - Urgency

  if (rowAry[0][colNoUid - 1] != colNameUid) return false;         // 9 - I - Info(Id)
  if (rowAry[0][colNoReport - 1] != colNameReport) return false;       // 10 - J - Report

  // the below cell is used for bulkupdate status checking
  // if (rowAry[0][colNoExtraDisplay1 - 1] != "Extra1") return false;       // 11 - K - Extra Display  
  if (rowAry[0][colNoDetails - 1] != colNameDetails) return false; // 12 - L - Details

  return true;
}


function createMissingColumnsTemp3()
{  

  const sheetsAry = getArrayOfAllScheduleSheets();

  for (let index = 0; index < sheetsAry.length; index++) {
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetsAry[index]);
    sheet.activate();
    LoggerLog("Processing: " + sheet.getName());
    sheet.getRange("I1").setValue(colNameUid);
    // sheet.getRange("A1").setValue(colNametSNo);
    // sheet.getRange("K1").setValue(colNoUid);
    mySleep(2);
    continue;

    // sheet.insertColumnAfter(10);
    // sheet.insertColumnAfter(11);
    // sheet.getRange("K1").setValue("Extra1");
    // sheet.getRange("L1").setValue("Details");
    
    sheet.insertColumnAfter(12);
    sheet.getRange("M1").setValue("Tags");  // 13  

    sheet.setColumnWidth(colNoUid, 10);
    sheet.setColumnWidth(colNoUrgency, 10);
    sheet.setColumnWidth(colNoUid, 50);
    sheet.setColumnWidth(colNoDetails, 50);
    sheet.setColumnWidth(colNoTags, 50);

    try{
      sheet.deleteColumn(14);  
      sheet.deleteColumn(15);
      sheet.deleteColumn(16);      
    }
    catch{}
  }
}

function IsDayTimeNow()
{
  var timeNow = new Date();
  //timeNow.setHours(04, 15, 59); //test

  var startTime = new Date();
  startTime.setHours(9, 30, 0);

  var endTime = new Date();
  endTime.setHours(21, 30, 0);

  var isDayTime = timeNow > startTime && timeNow < endTime
  Logger.log("isDayTime: " + isDayTime);
  return isDayTime;

  Logger.log("timeNow: " + timeNow);
  Logger.log("startTime: " + startTime);
  Logger.log("endTime: " + endTime);
}

function SaveExecutedTime(time){
  //var timeNow = SpreadsheetApp.getActive().getRangeByName(RangeTimeNow).getValue();
  //var timeNow = SpreadsheetApp.getActive().getRangeByName(RangeTimeLastExecuted).getValue();
  SpreadsheetApp.getActive().getRangeByName(RangeTimeLastExecuted).setValue(time);
}

function IfExecutedTimeAboveMinutes(mins){
  var timeNow = SpreadsheetApp.getActive().getRangeByName(RangeTimeNow).getValue();
  var timeLast = SpreadsheetApp.getActive().getRangeByName(RangeTimeLastExecuted).getValue();
  // Logger.log("Diff (> 60000 ?): "+ (timeNow - timeLast) ); //60000 = 1 minute
  var isAbove1min = (timeNow - timeLast) > (60000 * mins);  //60000 = 1 minute
  Logger.log("IfExecutedTimeAbove" + mins + "Minutes: (is >" + (60000 * mins) + "ms)" + isAbove1min); //60000 = 1 minute  
  return isAbove1min;
}

function LoggerLog(text){
  Logger.log(text);  
}

function PrintTimeNow(log){
  if(IsBulkUpdateDoingNow)return;

  var now = new Date();
  if(isNaN(oldtime)) NoteDownThisTimeAsOldTime();
  var bDay = oldtime; // new Date(2023, 01, 28);
  var elapsedT = now - bDay; // in ms    
  oldtime = now;
  LoggerLog("elapsedT: = " + elapsedT + ',  ' + log);
}

var oldtime = new Date();

function NoteDownThisTimeAsOldTime(){
  if(IsBulkUpdateDoingNow)return;
  oldtime = new Date();
}


Date.prototype.addDays = function (days) {
    const date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
};

Date.prototype.subDays = function (days) {
    const date = new Date(this.valueOf());
    date.setDate(date.getDate() - days);
    return date;
};

function addYears(date, years) {  
  date.setFullYear(date.getFullYear() + 1);
  return date;
}

function addMinutes(date, minutes) 
{ 
  date.setMinutes(date.getMinutes() + minutes);
  return date;
}

function addHours(date, hours) 
{ 
  date.setHours(date.getHours() + hours);
  return date;
}

function addDays(date, days) 
{   
  // var date = new Date(this.valueOf());
  date.setDate(date.getDate() + days);
  return date;

  // date.setDays(date.getDays() + days);
  // return date;
}


/*
Date.prototype.addDays = function(date, days) {

    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}
*/


function DisplayFinishedWorks()
{   
  LoggerLog("DisplayFinishedWorks(), Version 6");
  // var sheet = SpreadsheetApp.getActive().getSheetByName('Farook');
  // sheet.activate();

  var sheet = SpreadsheetApp.getActiveSheet();  
  var sheetName = sheet.getName();

  // var sheetCorrect = IsItScheduleSheetCheck2(sheetName);  
  // if (!sheetCorrect) return; 

  var sheetOk = IsItScheduleSheetCheck1And2(sheetName);
  if (!sheetOk) return;

  // var sheetOk = IsItScheduleSheetCheck1(sheetName);    
  var sheetOk = IsItScheduleSheetCheck1And2(sheetName);
  if (!sheetOk){ 
    // showAlert("Wrong sheet");
    return; 
  }

  var sheetWorksDone = SpreadsheetApp.getActive().getSheetByName('WorksDone');
  if (sheetWorksDone == null){
    var sheetWorksDone = SpreadsheetApp.getActive().insertSheet();
    sheetWorksDone.setName("WorksDone");
  }
  sheetWorksDone.showSheet();
  var temp3 = sheetWorksDone.getRange("A1").getValue();
  var daysback = GetNumberIfOrNull(temp3, 0);
  
  sheetWorksDone.getDataRange().clearContent();
  sheetWorksDone.getRange("A1").setValue(daysback.toString());

  // return;
  var rowPaste = 2;
  var dt = new Date();    
  dt = addDays(dt, daysback);
  var nt = dt.getDate()
  //var colNoOfTodayWork = colNoReport + nt;
  var sheetName = sheet.getName();
  if (sheetName == "WorksDone") return;
  var totalRows = sheet.getMaxRows();  
    
  sheetWorksDone.getRange(2, 1, 100, 5).clearContent();

  Logger.log("totalRows: " + totalRows);
  Logger.log("colNoReport: " + colNoReport);
  Logger.log("dt: " + convertDateShortWithHourMinutes(dt));    
  var todayYes = convertDateShort(dt)  + " Yes";      
  Logger.log("todayYes: " + todayYes.toString());      
  Logger.log("nt: " + nt.toString());
  Logger.log("sheetWorksDone: " + sheetWorksDone.getName());
  Logger.log("------------------------------------------");
  Logger.log("");
  // return;

  var data = sheet.getDataRange().getValues();
  sheetWorksDone.activate();
  sheetWorksDone.getRange('B1').activate();

  for (let rowNo = 2; rowNo < totalRows; rowNo++) 
  {
    //var rangeDue = sheet.getRange(rowNo, colNoDue);   
    //var dueTime2 = rangeDue.getValue();
    
    if(rowNo >= data.length) break;
    var status = data[rowNo-1][colNoStatus-1];
    if(status == 'Disabled'){
      continue;
    }

    var due = data[rowNo-1][colNoDue-1];
    dueTime2 = due;

    var rangeReport = sheet.getRange(rowNo, colNoReport);      
    var rangeWorkName = sheet.getRange(rowNo, colNoWork);

    var workName = rangeWorkName.getValue();
    if (workName == null || workName.toString().length<1){continue;}
    
    var allDaysReport = rangeReport.getValue();
    var firstLine = allDaysReport.split('\n')[0];

    var log = `Row:${rowNo}, Due:${convertDateShortWithHourMinutes(dueTime2)} - ${firstLine} - ${workName.toString().substring(0,15)}`;
    Logger.log(log);


    var rangeRepeat = sheet.getRange(rowNo, colNoRepeat);
    var valRepeat = rangeRepeat.getValue();

    firstLine = allDaysReport;

    if(rowNo == 3){
      LoggerLog("row 3");
    }

    if (allDaysReport.includes(todayYes))
    {
      sheetWorksDone.getRange(rowPaste, 2).setValue(workName);     
      sheetWorksDone.getRange(rowPaste, 3).setValue("Yes");   
      sheetWorksDone.getRange(rowPaste, 4).setValue(valRepeat);    
      rowPaste = rowPaste + 1;     
    }
    else // if (firstLine.includes("No"))
    {
      if(status == 'Over'){
        // It means, work was done many days ago, 
        // so there is no 'today Yes' in report column
        // Since it is 'over', work already done, so don't show it as, 'not done' or 'done' today.
        continue;
      }

      if(dt > due){
        sheetWorksDone.getRange(rowPaste, 2).setValue(workName);     
        sheetWorksDone.getRange(rowPaste, 3).setValue("No");  
        sheetWorksDone.getRange(rowPaste, 4).setValue(valRepeat);  
        //sheetWorksDone.getRange(rowPaste, 5).setValue("Due");  
        rowPaste = rowPaste + 1;   
      }
      LoggerLog("Not due yet");
    }
    //if (rowPaste>25) break;
  }

  sheetWorksDone.activate();
  sheetWorksDone.sort(3, false);
  
  //Good Code to add row gap between Done & NotDone works
  var valprev = sheetWorksDone.getRange(2, 3).getValue();
  var totalRows2 = sheetWorksDone.getDataRange().getNumRows();
  for (let rowy = 2; rowy < totalRows2; rowy++) {
    var valnow = sheetWorksDone.getRange(rowy, 3).getValue();
    if(valnow != valprev){
      sheetWorksDone.insertRowBefore(rowy);
      var valprev = valnow;
    }
  }

  /*
  //Good Code to add row gap between repeat
  var valprev = "";
  var totalRows2 = sheetWorksDone.getDataRange().getNumRows();
  for (let rowy = 2; rowy < totalRows2; rowy++) {
    var valnow = sheetWorksDone.getRange(rowy, 4).getValue();
    if(valnow != valprev){
      sheetWorksDone.insertRowBefore(rowy);
      var valprev = valnow;
    }
  }*/
  
  sheetWorksDone.getRange('B1').setValue(sheet.getName());
  // sheetWorksDone.getRange('B' + (rowPaste +2).toString()).activate();
  sheetWorksDone.getRange('B1').activate();
}


function setValueInCell(sheet, row, col, value){
  // not used now, but good
  sheet.getRange(row, col).setValue(value);
}

function addMonths(date, months) {
    var d = date.getDate();
    date.setMonth(date.getMonth() + +months);
    if (date.getDate() != d) {
      date.setDate(0);
    }
    return date;
}


function addNewJob(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var totalRows = sheet.getMaxRows();
  Logger.log("showAllWorks, totalRows: " + totalRows.toString());
  var rangeToUnHide = sheet.getRange(2, 1, totalRows-1);   
  sheet.unhideRow(rangeToUnHide);   
}


function showAllWorks(){
   
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();
  if (sheetName == "Base") return;

  var totalRows = sheet.getMaxRows();
  Logger.log("showAllWorks, totalRows: " + totalRows.toString());
  var rangeToUnHide = sheet.getRange(2, 1, totalRows-1);   
  sheet.unhideRow(rangeToUnHide);  
  
  // SortColumnByJobFrequency(sheet);    
  // Unmerge all cells in sheet: SpreadsheetApp.getActive().getActiveSheet().getDataRange().breakApart(); – 
  // clearformatting to unmerge
  // if(IsExtraDisplayEnabled) sheet.getRange("K:K").clearFormat();

  SortColumnsBy462(sheet);   
  OnlySetColWidth(sheet);

  setShowAll();
  goToCellB2(sheet);

  // mySleep(60);
  // pauseExecution();
  // Utilities.sleep(8000);

  return;
  mySleep(5);

}

const JobStatusAddress = "A1";
function setShowAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // sheet.getRange(JobStatusAddress).setValue("Job-show");
  sheet.getRange(JobStatusAddress).setNote("Job-show");
}
function setHideDone() {  
  var sheet = SpreadsheetApp.getActiveSheet();
  // sheet.getRange(JobStatusAddress).setValue("Job-hide");
  sheet.getRange(JobStatusAddress).setNote("Job-hide");
}
function getShowAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // return sheet.getRange(JobStatusAddress).getValue() == "Job-show";
  retval = sheet.getRange(JobStatusAddress).getNote() == "Job-show";
  LoggerLog(`getShowAll(): ${retval}`);
  return retval;
}

function getHideDone() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // return sheet.getRange(JobStatusAddress).getValue() == "Job-hide";
  return sheet.getRange(JobStatusAddress).getNote() == "Job-hide";
}


function goToCellB2(sheet){ 
  sheet.getRange('B2').activate();
}


function FormatResetColorAllRows(){
 //  Utilities.sleep(2000);
  var sheet = SpreadsheetApp.getActiveSheet();
  var totalRows = sheet.getMaxRows();
  var totalCols = sheet.getMaxColumns();
  Logger.log("showAllWorks, totalRows: " + totalRows.toString());
  var rangeToUnHide = sheet.getRange(2, 1, totalRows-1, totalCols);   
  rangeToUnHide.setBackground("white");
}


 



function testDbDate1()
{  
  var sheet = SpreadsheetApp.getActiveSheet();  
  var DbDate1 = sheet.getRange("Db!A3");
  var res = DbDate1.getValue();
  Logger.log("DbDate1: " + convertDate(res));
}

function convertDate(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  return [pad(d.getDate()), pad(d.getMonth()+1), d.getFullYear()].join('/')
}

function convertDateShort(inputFormat) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  // return [pad(d.getDate()), pad(d.getMonth()+1), d.getYear()].join('.')
  return  Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "dd.MM.yy")
}

function convertDateShortWithHourMinutes(inputFormat) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat)
  // return [pad(d.getDate()), pad(d.getMonth()+1), d.getYear()].join('.')
  return  Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "dd.MM.yy hh:mm a")
}

function FormatFirstRow(sheet, busy) {

  var rangeFirstRow = sheet.getRange('1:1');
  sheet.setFrozenRows(1);

  // sheet.getRange('1:1').activate();

  if(busy){
    rangeFirstRow
    .setBackground('#8a0032')
    .setFontColor('#ffffff')  
    .setHorizontalAlignment('left')
    .setFontSize(16);
    sheet.setRowHeightsForced(1, 1, 46);
  }
  else{
    rangeFirstRow
    .setBackground('#6aa84f')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('left')
    .setFontSize(16);
    sheet.setRowHeightsForced(1, 1, 36);
  }
}

function IsItScheduleSheetCheck1And2(activeSheetName){

  var check1 =  IsItScheduleSheetCheck1(activeSheetName);
  var check2 =  IsItScheduleSheetCheck2(activeSheetName);
  return (check1 && check2)
}

function IsItScheduleSheetCheck1(activeSheetName){

  var empNames = SpreadsheetApp.getActive().getRangeByName(ScheduleSheetsRangeInDB).getDisplayValues();   
  LoggerLog("IsItScheduleSheetCheck1(): activeSheetName= " + activeSheetName);

  for (let ind = 0; ind < empNames.length; ind++) {

    var emp = empNames[ind];
    // Logger.log("emp: " + empNames[ind]);
    if (emp == null || emp == "") break;
    if (emp == activeSheetName) {
      LoggerLog("IsItScheduleSheetCheck1(): sheetOk: true");
      return true;
    }
  }
  
  LoggerLog("IsItScheduleSheetCheck1(): sheetOk: false");
  return false;
}


function IsItScheduleSheet2(activeSheet){  
  // DROPPED - REASON BELOW

  // var sheetOk2 = IsItScheduleSheet2(activeSheet);  
  // LoggerLog("sheetOk2: " + sheetOk2);
  // PrintTimeNow('onEdit-Code-103');
  // ------------------

  // E3 created this function to reduce execution time for
  // finding 'the sheetname exists in to process for schedule or not'
  // by having a text in K1 range. But it take 130ms
  // the other method (IsItScheduleSheet) takes 260ms. so E3 dropped this method.

  return activeSheet.getRange("K1").getValue().includes('vdo');
}


function getDaysToPostponeFromText(text, dateToPostpone){

   var days = -1;

   // Postpone x mins
   if (text.includes("mins")) {           
      var minutesStr = text.replace(/\D/g,'').trim();
      var minutes = Number(minutesStr);
      Logger.log("minutes: " + minutes);    
      dateToPostpone = addMinutes(dateToPostpone, minutes);
      Logger.log("dateToPostpone: " + dateToPostpone);    
      return dateToPostpone;     
   }

   // Postpone x hours
   if (text.includes("hour")) {           
      var hoursStr = text.replace(/\D/g,'').trim();
      var hours = Number(hoursStr);
      Logger.log("hours: " + hours);    
      dateToPostpone = addHours(dateToPostpone, hours);
      Logger.log("dateToPostpone: " + dateToPostpone);    
      return dateToPostpone;     
   }

   // Postpone x days
   if (text.includes("day")) {           
      var daysStr = text.replace(/\D/g,'').trim();
      var days = Number(daysStr);
      Logger.log("days: " + days);    
      dateToPostpone = addDays(dateToPostpone, days);
      Logger.log("dateToPostpone: " + dateToPostpone);    
      return dateToPostpone;     
   }

   // Postpone x weeks
   if (text.includes("week")) {           
      var weeksStr = text.replace(/\D/g,'').trim();
      var weeks = Number(weeksStr);
      Logger.log("weeks: " + weeks);    
      days = weeks * 7;
      dateToPostpone = addDays(dateToPostpone, days);
      Logger.log("dateToPostpone: " + dateToPostpone);    
      return dateToPostpone;     
   }

   // Postpone x months
   if (text.includes("month")) {           
      var monthsStr = text.replace(/\D/g,'').trim();
      var months = Number(monthsStr);
      dateToPostpone = addMonths(dateToPostpone, months);
      return dateToPostpone;     
   }

   // Postpone x years
   if (text.includes("year")) {           
      var yearsStr = text.replace(/\D/g,'').trim();
      var years = Number(yearsStr);
      Logger.log("years: " + years);    
      dateToPostpone = addYears(dateToPostpone, years);
      Logger.log("dateToPostpone: " + dateToPostpone);    
      return dateToPostpone;     
   }

   if (days != -1){
      dateToPostpone = dateToPostpone.addDays(days);
      return dateToPostpone;
   }
   // if (days == -1) it is either month or year, so continue;

  showAlert("From getDaysToPostponeFromText, return value: Null");
  return null;
}

function isNumeric2(str) {
  if (typeof str != "string") return false // we only process strings!  
  return !isNaN(str) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
         !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
}

function isNumeric3(value) {
  return !isNaN(parseFloat(value)) && isFinite(value);
}

function isNumeric1(substring1) {
    if (!isNaN(parseFloat(substring1)) && isFinite(substring1)) {
        return true;
    } else {
        return false;
    }
}

function getLocation() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rangeActive = sheet.getActiveRange();
  var rowNow = rangeActive.getRow();
  var rngCur = sheet.getRange(rowNow, colNoWork);   
  var val1 = rngCur.getValue();
  Logger.log("rngCur: " + val1);

  return {
    sheet: val1,
    range: 'done'
  };
}

function DelCheckDayOfToday(){
  var dt = new Date();
  var nt = dt.getDate()  
  Logger.log("nt: " + nt);
  Logger.log("dt: " + dt);
  var colNoOfTodayWork = colNoStatus + nt;

  Logger.log("colNoOfTodayWork: " + colNoOfTodayWork);
  var sheet = SpreadsheetApp.getActiveSheet();
  var rangeOfTodayWork = sheet.getRange(2, colNoOfTodayWork);   
  Logger.log("rangeOfTodayWork: " + rangeOfTodayWork.getA1Notation());

    return;
}


function showAdminSidebar() {
 var widget = HtmlService.createHtmlOutputFromFile("Adminpage.html");
 //var widget = HtmlService.createHtmlOutput("<h1>Sidebar</h1>");
 SpreadsheetApp.getUi().showSidebar(widget);
}

function CallRefreshWorks(){
  if (!IsDayTimeNow()) return;
  refresh2(0, WorkHideDoneWorks);
}



function HideDoneWorks(){ 

  // setHideDone();
  // refresh2(0, WorkHideDoneWorksNow);   
  // var sheet = SpreadsheetApp.getActiveSheet();
  // sheet.setColumnWidth(colNoSNo, 1); 
  // sheet.setColumnWidth(colNoProperTime, 1); 
  // sheet.setColumnWidth(colNoDue, 130); 
  // goToCellB2(sheet);
  // return;
  // -------------------------------------------

  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetName = sheet.getName();
  if (sheetName == "Base") return;
  
  var correctSheet = IsItScheduleSheetCheck1And2(sheetName);
  LoggerLog(`HideDoneWorks():sheetCorrect: ${correctSheet}`);
  if (correctSheet == false) return;

  AutoUpdateStatusOfSheet(sheetName);
  
  setHideDone();

  // acutally the following line will not hide anymore!
  // since logic changed.
  // refresh2(0, WorkHideDoneWorksNow);  

  OnlyHideDoneRows();

  // HideColumnsAndSetWidths(sheet);  
  goToCellB2(sheet);  
}


function PageFormatting(){   
  //Main function that formates and rectify all errors.
  LoggerLog("PageFormatting():");

  var sheet = SpreadsheetApp.getActiveSheet();  
  DoPageFormatting(sheet);
  goToCellB2(sheet);
}

function DoPageFormatting(sheet){

  var sheetName = sheet.getName();
  // var sheetCorrect = IsItScheduleSheetCheck2(sheetName);  
  // if (!sheetCorrect) return;  
  // var sheetOk = IsItScheduleSheetCheck1(sheetName);    
  // if (!sheetOk) return;
  
  var sheetOk = IsItScheduleSheetCheck1And2(sheetName);
  if (!sheetOk) return;

  SetTitle(sheet);  
  PutUidForAllJobs(sheet);
  FormatFirstColumn(sheet); 
  FormatFirstRow(sheet, false);
  FormatDateFull(sheet);
  SetValidationForStatus(sheet);
  SetValidationForFrequency(sheet);  //Repeat
  SetValidationForEmps(sheet);
  SetValidationForUrgency(sheet);
  // SortColumnByJobFrequency(sheet);
  
  SortColumnsBy462(sheet);  
  HideColumnsAndSetWidths(sheet);
  FormatCellAlignment(sheet);
  FormatRowHeights(sheet);
  SetTitle(sheet);

}



function SetValidationForStatus(sheet){

  var totalRows = sheet.getMaxRows();
  var range = sheet.getRange(StartRow, colNoStatus, totalRows - (StartRow - 1), 1);

  range.setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(sheet.getRange('Db!$D:$D'), true)
  .build());
}



function SetValidationForFrequency(sheet){

  var totalRows = sheet.getMaxRows();
  var range = sheet.getRange(StartRow, colNoRepeat, totalRows - (StartRow - 1), 1);

  range.setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(sheet.getRange('Db!$E:$E'), true)
  .build());
}


function SetValidationForEmps(sheet){  

  var totalRows = sheet.getMaxRows();
  var range = sheet.getRange(StartRow, colNoStaff, totalRows - (StartRow - 1), 1);

  range.setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(sheet.getRange('Db!$C:$C'), true)
  .build());
}


function SetValidationForUrgency(sheet){

  var totalRows = sheet.getMaxRows();
  var range = sheet.getRange(StartRow, colNoUrgency, totalRows - (StartRow - 1), 1);

  range.setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(sheet.getRange('Db!$F:$F'), true)
  .build());
}

function HideColumns1And3() {
  var spreadsheet = SpreadsheetApp.getActive();
  //spreadsheet.getRangeList(['A:A', 'C:C']).activate();
  spreadsheet.getRangeList(['A:A']).activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  // spreadsheet.getRange('B:B').activate();
  spreadsheet.getRange('B1').activate();
};


function HideColumnsAndSetWidths(sheet) {  
  OnlySetColWidth(sheet);
  OnlyHideColumns(sheet);  
  sheet.getRange('B1').activate();
}

function OnlyHideColumns(sheet){
  // var range = sheet.getRange("A1");
  // sheet.unhideColumn(range);

  sheet.hideColumns(colNoSNo);
  sheet.hideColumns(colNoProperTime);
  sheet.hideColumns(colNoExtraDisplay1);
  sheet.hideColumns(colNoUrgency);
  sheet.hideColumns(colNoUid);  
  sheet.hideColumns(colNoTags);
}

function OnlySetColWidth(sheet){
  sheet.setColumnWidth(colNoSNo, 70);
  sheet.setColumnWidth(colNoWork, 316); //Title

  sheet.setColumnWidth(colNoProperTime, 263); //Job Time 263
  sheet.setColumnWidth(colNoDue, 263); //Postponed Time  

  // sheet.setColumnWidth(colNoUrgency, 1);
  // sheet.setColumnWidth(colNoExtraDisplay1, 1);
  // sheet.setColumnWidth(colNoUid, 1);
  // sheet.setColumnWidth(colNoTags, 1);

  sheet.setColumnWidth(colNoStatus, 75);  //Status
  sheet.setColumnWidth(colNoProperTime, 70); //Job Time
  sheet.setColumnWidth(colNoDue, 130); //Postponed Time
  sheet.setColumnWidth(colNoRepeat, 130); //Repeat
  sheet.setColumnWidth(colNoStaff, 60); //Staff
  sheet.setColumnWidth(colNoExtraDisplay1, 40); //Staff
  sheet.setColumnWidth(colNoUrgency, 40); //Urgency
  sheet.setColumnWidth(colNoUid, 40);  //Extra1 forSorting based on Repeat. Now we use to write, Info/Uid
  sheet.setColumnWidth(colNoReport, 150); //Report
  sheet.setColumnWidth(colNoDetails, 300);   
  sheet.setColumnWidth(colNoTags, 60); //Staff
  
}

function FormatFirstColumn(Actsheet){
  Actsheet.getRange("A:A").setFontColor("#e8e8e8");
  Actsheet.getRange("A1").setFontColor("#60a667");
}

function FormatCellAlignment(sheet) {

  var rows = sheet.getMaxRows();
  var cols = sheet.getMaxColumns();

  // Column DueDate text color is grey
  sheet.getRange(2,colNoDue, rows -1, 1).setFontColor("#dbdbdb");

  // column Report text color is grey
  sheet.getRange(2,colNoReport, rows -1, 1).setFontColor("#dbdbdb");
  
  // column Report text color is grey
  sheet.getRange(2,colNoRepeat, rows -1, colNoDetails - colNoRepeat).setFontColor("#dbdbdb");

  // Other cells font size, alignments, etc..
  var rng = sheet.getRange(1, 1, rows, cols );
  rng.setFontSize(16)
  .setFontFamily('Verdana')
  .setHorizontalAlignment('left')
  .setVerticalAlignment('top');
  
  rng.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); 
  // sheet.getRange('B1').activate();

  
}

function FormatRowHeights(sheet) {
  var totalRows = sheet.getMaxRows();  
  sheet.setRowHeightsForced(1, totalRows, 30);  
  // sheet.getRange('B1').activate();
}
 

function atest2() 
{  
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.setRowHeightsForced(2, 3, 41);
};


function FormatDateFull(sheet) {
  LoggerLog("FormatDateFull(): SheetName: " + sheet.getName());  
  sheet.getRange('E:F').activate();
  sheet.getActiveRangeList().setNumberFormat('dd"."mm"."yyyy" "hh"."mm" "am/pm');
  
  sheet.getRange('J:J').activate();
  sheet.getActiveRangeList().setNumberFormat('@');   
}

 
function MoveRowToAnotherSheet(row) { 
  
var status = GetIsPausedForBulkUpdate();
if(status){LoggerLog("Current sheet is in BulkUpdate Mode, so exiting..."); return;}

//This function do 2 things.
//1. If we give commands like all/hide/del then it will do such things.
//2. If we select staff name then the row will be moved to corresponding staff page.

// var ui = SpreadsheetApp.getUi();  
var ss = SpreadsheetApp.getActiveSpreadsheet();  
var SourceSheet = SpreadsheetApp.getActiveSheet();  
var SourceSheetName = SourceSheet.getName();  
// var row = SourceSheet.getActiveRange().getRow();  //Works

var ValueInColReport = SourceSheet.getRange(row,colNoReport).getValue();  
var DestinationSheetName = SourceSheet.getActiveRange().getValue();     //Works
var date = Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yy");

//1. If we give commands like all/hide/del then it will do such things.
//------------------------------------------------------------------
if (DestinationSheetName == "") return;   

if (DestinationSheetName == "---------------") {
  var range = SourceSheet.getRange(row,colNoStaff);
  range.clearContent();
  return;   
}

//Add new rows at end. Bcoz mobiles view do not show menus, I write this function here too
//------------------------------------------------------------------
  if (DestinationSheetName == "Add 5 Rows") { 
  Add5row();
  var range = SourceSheet.getRange(row,colNoStaff);
  range.clearContent();
  }



//Show all rows including hidden. Useful in mobile view. Bcoz mobiles view do not show menus, I write this function here too
//------------------------------------------------------------------
if (DestinationSheetName == "Show All Works") { 
  showAllWorks();

  var range = SourceSheet.getRange(row,colNoStaff);
  range.clearContent();

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2').activate();
  return;
}


//Hide done works by typing hide in last empty cell. Bcoz mobiles view do not show menus, I write this function here too
//------------------------------------------------------------------
if (DestinationSheetName == "Hide Done Works") 
{   
  HideDoneWorks();    
  
  var range = SourceSheet.getRange(row,colNoStaff);
  range.clearContent();
  SourceSheet.getRange('B2').activate();
  return;
}

//Exit, if worksheet with that staff name not exists.
//--------------------------------------------------------------------
var DestinationSheet = ss.getSheetByName(DestinationSheetName);
if (!DestinationSheet) {      
    showAlert("Page " + DestinationSheetName + " not found. Either create page for him or move to someothers page.");
    return;
}

//Cut and paste the row want to move
//--------------------------------------------------------------------
var spreadsheet = SpreadsheetApp.getActive();
SourceSheet.getRange(row + ':' + row).activate();
SourceSheet.setCurrentCell(spreadsheet.getRange('B' + row));

var status = GetIsPausedForBulkUpdate(DestinationSheet);
if(status){LoggerLog(DestinationSheet.getName() + " is in BulkUpdate Mode, so exiting..."); return;}

//Get first empty row
//--------------------------------------------------------------------
ss.setActiveSheet(DestinationSheet, true);
LoggerLog("DestinationSheet: " + DestinationSheet.getName());
// return;

var range = DestinationSheet.getDataRange(); // may be unused
var values = DestinationSheet.getDataRange().getValues();
var FirstBlankRow = 0;
for (var FirstBlankRow=0; FirstBlankRow<values.length; FirstBlankRow++) {
  if (!values[FirstBlankRow].join("")) break;
}

PauseForBulkUpdate();
FirstBlankRow = FirstBlankRow + 1;
// DestinationSheet.getRange(FirstBlankRow + ':' + FirstBlankRow).activate();
DestinationSheet.getRange("B"+ FirstBlankRow).activate();
var sourceData = SourceSheet.getRange(row, 1, 1, colNoReport).getValues();
DestinationSheet.getRange(FirstBlankRow, 1, 1, colNoReport).setValues(sourceData);
DestinationSheet.getRange(FirstBlankRow, colNoStaff).setValue('');
SourceSheet.deleteRow(row);
ResumeForBulkUpdate();

// Alert user to add notes if any
// --------------------------------------------------------------------
showAlert('இப்போது move செய்த வேலையில்' + "\n" + '--------------------------' + "\n" + 'என்ன வேலை முடிந்துள்ளது' + "\n" + 'அடுத்து என்ன செய்யவேண்டும்' + "\n" + '--------------------------' + "\n" + 'போன்ற விபரங்கள் இருந்தால் Report column த்தில் F2 press செய்து, முதல் வரியில் சேர்க்கவும்.');
return;


//Paste at empty row
//--------------------------------------------------------------------
spreadsheet.getRange( SourceSheetName + "!" + row + ":" + row ).moveTo(spreadsheet.getActiveRange());

//Remove staff name
//--------------------------------------------------------------------
SpreadsheetApp.getActiveSheet().getRange(FirstBlankRow, colNoStaff).setValue('');
SpreadsheetApp.getActiveSheet().getRange(FirstBlankRow, colNoReport).setValue(date + ' from ' + SourceSheetName + "\n" + ValueInColReport);
//SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DestinationSheetName).getRange(RANGE).setValue('');

//Delete row from soucesheet.
//--------------------------------------------------------------------
var doc1 = SpreadsheetApp.getActiveSpreadsheet();
var sheet1 = doc1.getSheetByName(SourceSheetName); 
sheet1.deleteRow(row);   

};



function mySleep (sec)
{
  SpreadsheetApp.flush();
  Utilities.sleep(sec*1000);
  SpreadsheetApp.flush();
}

function GetActiveSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function PutUidForAllJobs(Actsheet){
  LoggerLog("PutUidForAllJobs():");

  if (Actsheet == null) return;
  // var sheetOk = IsItScheduleSheetCheck1(Actsheet.getName());  
  // if (!sheetOk) return;
  
  var sheetOk = IsItScheduleSheetCheck1And2(Actsheet.getName());
  if (!sheetOk) return;

  var firstEmptyRow = GetFirstEmptyRow();
  
  var UidRangeValues = Actsheet.getRange(1, colNoUid, firstEmptyRow - 1, 1).getValues();    
  var arrayLength = UidRangeValues.length;  
  LoggerLog(`PutUidForAllJobs(): arrayLength: ${arrayLength}`);  

  for (let row = 0; row < arrayLength; row++)
  {    
    LoggerLog(`PutUidForAllJobs(): Prev RowIndex: ${row},  Uid: ${UidRangeValues[row][0]}`);   

    oldUid = UidRangeValues[row];
    if(oldUid == ""){
      newUid = GetUniqueId();
      UidRangeValues[row][0] = newUid;
    }    
    LoggerLog(`PutUidForAllJobs(): RowIndex: ${row},  Uid: ${UidRangeValues[row][0]}`);   
  }

  // set at once
  Actsheet.getRange(1, colNoUid, firstEmptyRow - 1, 1).setValues(UidRangeValues);
}



function GetFirstEmptyRow(Actsheet){  
  //Get first empty row

  if (typeof Actsheet === 'undefined') { 
    var Actsheet = SpreadsheetApp.getActiveSheet();  
    // var Actsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  var values = Actsheet.getDataRange().getValues();
  var FirstBlankRow = 1;
  for (var FirstBlankRow=1; FirstBlankRow<values.length; FirstBlankRow++) {
    if (!values[FirstBlankRow].join("")) break;
  }
  
  if(FirstBlankRow == Actsheet.getMaxRows()){
    Actsheet.insertRowAfter(FirstBlankRow);
  }

  FirstBlankRow = FirstBlankRow + 1;
  LoggerLog(`GetFirstEmptyRow(): FirstBlankRow: ${FirstBlankRow}`);
  return FirstBlankRow;
}


function GetUniqueId(length = 12) {
    let result = '';
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    const charactersLength = characters.length;
    let counter = 0;
    while (counter < length) {
      result += characters.charAt(Math.floor(Math.random() * charactersLength));
      counter += 1;
    }
    return result;
}



function findCell(textToFind) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == textToFind) {
        return i+1;
        row = values[i][j+1];
        Logger.log(row);
        Logger.log(i); // This is your row number
      }
    }    
  }  
}



function UpdateFromAppSheetApp(){
  // This function is automatically run by a trigger for 10/30 minutes

  const rowLogStart = 2;
  const colLogId = 1;
  const colLogSheetName = 2;
  const colLogUid = 3;
  const colLogAction = 4;
  const colLogRowNo = 5;
  const colLogInfo1 = 6;
  const colLogInfo2 = 7;
  const colLogInfo3 = 8;
  
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("logs");  
  var id = logSheet.getRange(rowLogStart, colLogId, 1,1).getValue();
  if(id == "") return;
  var SheetNameToUpdate	= logSheet.getRange(rowLogStart, colLogSheetName, 1,1).getValue();
  // LoggerLog(`UpdateFromAppSheetApp(): UpdateStatus(): SheetName1: ${SheetNameToUpdate}`);    
  if(SheetNameToUpdate == "") return;

  var Uid1 = logSheet.getRange(rowLogStart, colLogUid, 1,1).getValue();
  if(Uid1 == "") return;
  var Info3 = logSheet.getRange(rowLogStart, colLogInfo3, 1,1).getValue();
  if(Info3 == "Done") return;
  var Action =  logSheet.getRange(rowLogStart, colLogAction, 1,1).getValue();
  if(Action == "") return;
  var editRowNo =  logSheet.getRange(rowLogStart, colLogRowNo, 1,1).getValue();  
  if(editRowNo < 2) return;
  LoggerLog(`UpdateFromAppSheetApp(): UpdateStatus(): SheetNameToUpdate: ${SheetNameToUpdate}`);  

  var editSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetNameToUpdate);  
  var Uid2 = editSheet.getRange(editRowNo, colNoUid, 1,1).getValue();
  var work = editSheet.getRange(editRowNo, colNoWork, 1,1).getValue();

  if(Uid1 == Uid2){      
    LoggerLog(`UpdateFromAppSheetApp(): UpdateStatus(): Uid: ${Uid2}`);  
    LoggerLog(`UpdateFromAppSheetApp(): UpdateStatus(): work: ${work}`); 
    logSheet.getRange(rowLogStart, colLogInfo3, 1,1).setValue("Done");    
    refresh2(editRowNo, WorkUpdate, editSheet)
    logSheet.deleteRow(rowLogStart);
    return;
  }

}


function ClearJobReport(sheet){ // Not used

  if (typeof sheet === 'undefined') { 
    var sheet = SpreadsheetApp.getActiveSheet();
  }
  CreateJobReport(sheet, true);
}

function getDayOfToday() {
  var today = new Date();
  var day = today.getDate();
  return day;
}

function createColumnsOnRightMost(sheet, numColumnsToAdd = 1) {
  if(numColumnsToAdd < 1)
    return;
  var currentNumColumns = sheet.getLastColumn();
  sheet.insertColumnsAfter(currentNumColumns, numColumnsToAdd);
}

function CreateJobReportAndPdf(){
  UseSelectedCellAsMonth = false;
  var success = CreateJobReport();
  // var dam1 = getCurrentDateAsmmyy();
  var dam2 = getFileNameDatePartForPDF();
  if(success) generatePdf(dam2);
}


function CreateJobReportAndPdfForMonth(){
  UseSelectedCellAsMonth = true;
  var success = CreateJobReport();
  var dam1 = getOtherMonthDateAsmmyy();
  var dam2 = getFileNameDatePartForPDF();
  var dam = dam2 + ` (of ${dam1})`;
  if(success) generatePdf(dam);
}

function getFileNameDatePartForPDF(){  
  var datePart = Utilities.formatDate(new Date(), "GMT+05:30", "dd.MM.yy - hh.mm");
  Logger.log(`datePart= ${datePart}`); 
  return datePart;

  // const currentDate = new Date();
  // Logger.log(`currentDate= ${currentDate}`); 
}

function CreateJobReport(sheet, justClear = false){

  if (typeof sheet === 'undefined') { 
    var sheet = SpreadsheetApp.getActiveSheet();
  }
  Logger.log(`sheet= ${sheet.getSheetName()}`); 
  
  var days = 31;
  var works = 5;
  var works = sheet.getMaxRows();
  var totalCols = sheet.getMaxColumns();

  // Logger.log(`totalRows= ${totalRows}`);
  Logger.log(`totalCols= ${totalCols}`);

  if(totalCols < colNoJobReportDay1 + 31){
    var colsToCreate =  colNoJobReportDay1 + 31 - totalCols;
    Logger.log(`totalCols= ${totalCols}, so creating ${colsToCreate} columns...`);
    createColumnsOnRightMost(sheet, colsToCreate);
  }else{
    // return;
  }

  // sheet.getRange(1, colNoJobReportDay1).setValue(''); 
  sheet.getRange(1, colNoJobReportDay1-1).setValue('JR');  
  // sheet.hideColumns(colNoTags+1, colNoJobReportDay1-2-colNoTags);
  sheet.hideColumns(colNoUrgency, colNoJobReportDay1-colNoUrgency);
  sheet.unhideColumn(sheet.getRange(1,colNoReport));
  

  // unhide rows
  var rRows = sheet.getRange("A:A");
  sheet.unhideRow(rRows);

  // CREATE AN EMPTY 2D ARRAY
  var values = createEmptyArray(works, days);

  // CREATE FIRST ROW WITH DATES 1 TO 31 
  for (var day = 0; day < days; day++) {      
    values[0][day] = (day + 1).toString();    
  }

  // SET COLUMN WIDTHS FOR JOBREPORTS
  var totalCols = sheet.getMaxColumns();
  Logger.log(`totalCols2= ${totalCols}`);
  var cols2 = (totalCols-colNoTags);
  // var cols3 = (cols2 > 0)? cols2: 1;

  if(cols2 > 1){
    sheet.setColumnWidths(colNoTags+1, cols2-1, 30);
  }

  // CLEAR ALL AREA, AND SET DATES IN FIRST ROW
  sheet.getRange(1, colNoJobReportDay1, values.length, values[0].length).setValues(values);  
  // ------------------------------------------------------------------------
  

  // Set the font color to black (hex code #000000)
  var rangeb = sheet.getRange(2,colNoJobReportDay1,works-1, totalCols-colNoJobReportDay1+1); // Change this to your desired range
  rangeb.setFontColor("#000000");
  rangeb.setBackground("#ffffff");
  rangeb.setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);

  // ------------------------------------------------------------------------

  // Clearing backgrounds is over at this point, so exit if required.
  if(justClear)return;



  // var dayno = getDayOfToday();
  // var dayno = sheet.getActiveCell().getValue();
  // var dayno = sheet.getRange(sheet.getCurrentCell().getRow() + 1,1).getValue();

  var currentCell = sheet.getSelection().getCurrentCell();
  if (currentCell !== null) {
    var dayno = currentCell.getValue();
  }

  LoggerLog(`CreateJobReport: dayno:${dayno}`);
  if((isNumeric1(dayno) === false) || (dayno == '')){
    dayno = getDayOfToday();
  } 

  LoggerLog(`CreateJobReport: dayno:${dayno}`);
  dayno = dayno.toString().trim();

  if((isNumeric1(dayno) === false) || (dayno == '')){
    LoggerLog(`dayno is not numeric`);
    showToast('Select any date!');
    return;
  }

    
  // showToast(`dayno:${dayno}`);
  // mySleep(4);
  // showToast(`values.length:${values.length}`);

  // ------------------------------------------------------------------------

  
  // GET JOBREPORT DATA IN ARRAY:valuesReport
  var rangeReport = sheet.getRange(1, colNoReport, works, 1);  
  var valuesReport = rangeReport.getValues();

  // rangeReport.activate();
  // Logger.log("rangeReport= " + rangeReport.getA1Notation());
  //  LoggerPrintArrayData(valuesReport, 'valuesReport')
  //  return;

  var dt = new Date();
  var month = (dt.getMonth() + 1).toString().padStart(2, '0'); // Adding 1 to month since it's zero-based
  var year = dt.getFullYear().toString().slice(-2);
  // LoggerLog(`Month:${month}, Year:${year}`);

  if(UseSelectedCellAsMonth){

    LoggerLog(`CreateJobReport: dayno:${dayno}`);
    if(dayno > 12){
      LoggerLog(`Month should be below 13`);
      showToast(`${dayno}th month?  For Other month reports, month should be below 13!`);
      return;
    }
    month = dayno.toString().padStart(2, '0');
  }  
  
  var backColor = UseSelectedCellAsMonth?  "#ffccee" : "#cceeff"; // Rose(Other month) : Blue(current Month)
  // Set the background color of today's column (hex code #000000)
  // var ranget = sheet.getRange(2,colNoJobReportDay1,works-1, totalCols-colNoJobReportDay1); 

  LoggerLog(`CreateJobReport, 3069: dayno:${dayno}`);

  dayno = stringToInt(dayno);
  var ranget = sheet.getRange(2, colNoJobReportDay1 + dayno - 1, works - 1, 1);
  ranget.setBackground(backColor);


  // MAIN SECTION - LOGS ALL JOB REPORTS FOR DATES 1 TO 31
  // WORK=0 IS THE FIRST ROW, SO LEAVE IT
  for (var work = 1; work < works; work++) { 

    var reportData = valuesReport[work][0];
    if(reportData == '') continue;
    // LoggerLog("reportData: " + reportData); 
    // Logger.log(`work: ${work}, reportData: ${reportData}`);

    for (var day = 0; day < days; day++) {
      
      var dd = (day<9)? "0" + (day+1).toString() : (day+1).toString();
      
      // var regexPattern = /(\b04\.09\.23\ +)(\w+)/
      // THE SAME ABOVE WRITTEN AS BELOW, COMPATIBLE FOR ARGUMENT PASSING
      var regexPattern = "(\\b" + dd + "\\." + month + "\\." + year + "\\s+)(\\w+)"
       
      // LoggerLog("regexPattern: " + regexPattern);
      var regExp = new RegExp(regexPattern, "i"); // g=global, i=case-insensitive
      var anyWordAfterDate = false;        

      while (match = regExp.exec(reportData)) {      
        // Logger.log(`match[0]: ${match[0]}`); // 0th element contains full match
        // Logger.log(`match[1]: ${match[1]}`); // 1st element contains 1st match
        // Logger.log(`match[2]: ${match[2]}`); // 2nd element contains 2nd match
        
        var wordAfterDate = match[2];
        if(wordAfterDate != null && wordAfterDate != 'undefined' && 
          wordAfterDate != 'No' && wordAfterDate != 'Created' && wordAfterDate.length > 2){
          // means somework done like;  '10.10.23 called hasan' or  '10.10.23 Yes'
          anyWordAfterDate = true; 
        }

        // No need to break, since we didn't search global in RegExp
        // only 1 result will be there, be on safer side, also for future
        break; 
      }

      // old code, just checks whether it includes
      // var isMatch = reportData.includes(todayYes);
      // dont remove the above code, it may be useful if regex not works!
      
      if (anyWordAfterDate) {    
        // values[work][day] = '✅'; // Not shows in PDF
        // values[work][day] = '✔';  // Very tick
        var status = (wordAfterDate == "Worked")? 'W': '✓'; 
        // values[work][day] = '✓';
        values[work][day] = status;
      } 
      else {
        // values[work][day] = '❌';//  ⭕ not displayed in exported PDF
        values[work][day] = '';
      }

      // ⭕
    }
  }

  // Logger.log("values: " + values);
  // LoggerPrintArrayData(valuesReport, 'valuesReport')

  sheet.getRange(1, colNoJobReportDay1, values.length, values[0].length).setValues(values); 

  SortByFrequency();
  // return; // remove it farook e3

  return true; 
}

function stringToInt(input) {
   return parseInt(input, 10);

  if (typeof input === 'string' && input !== '') {
    return parseInt(input, 10);
  }
  return 10;
}


function createEmptyArray(rows, cols) {
  var emptyArray = [];
  
  for (var i = 0; i < rows; i++) {
    var row = [];
    for (var j = 0; j < cols; j++) {
      row.push('');
    }
    emptyArray.push(row);
  }
  
  return emptyArray;
}

function LoggerPrintArrayData(array, arrayName){  

  Logger.log(`${arrayName}: ${array}`);  
  Logger.log(`Dimentions: [${array.length}][${array[0].length}]`);   

  // Loop through the rows and columns and print the data
  for (var i = 0; i < array.length; i++) {
    for (var j = 0; j < array[i].length; j++) {
      Logger.log(`values[${i}][${j}]: ${array[i][j]}`);
      // Logger.log('Row ' + (i + 1) + ', Column ' + (j + 1) + ': ' + arrayToPrint[i][j]);
    }
  }
}


function E3Test(){
  Callcreateblobpdf();
}

function Callcreateblobpdf() {
   generatePdf();
}

function generatePdf(dateAndMonth) {
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSpreadsheet = SpreadsheetApp.getActive(); // Get active spreadsheet.
  var sheets = sourceSpreadsheet.getSheets(); // Get active sheet.
  var sheetName = sourceSpreadsheet.getActiveSheet().getName();
  var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
  var pdfName = sheetName + " - " + dateAndMonth + ".pdf"; // Set the output filename as SheetName.

  var folder = getDriveFolderByIdForPDF();  
  // var folder = getDriveFolderByName("Ssmc Job Reports");    
  if(folder == null){
    
    showToast('PDF folder is null!');
    return;
  }
  // DriveApp.getFoldersByName
  try {
    deleteFileByName(folder.getName(), pdfName);
    // return;
  }
  catch (e) {
      // handle the unsavoriness if needed
  }


  var theBlob = createblobpdf(sheetName, pdfName);  
  var newFile = folder.createFile(theBlob);

  // MUST FOR OTHERS TO DELETE AND CREATE A NEW FILE
  // var sharingPermission = DriveApp.Permission.EDIT;
  // newFile.setSharing(sharingPermission, DriveApp.Access.ANYONE);

  // Anyone with link can edit
  newFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT); 
  var durl = newFile.getDownloadUrl();
  openUrl1(durl);

  // var folder = DriveApp.createFolder('Shared Folder');
  // folder.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
}


function deleteFileByName(folderName, fileName ) {
  // Specify the folder where the file is located
  // var folderName = "Your Folder Name"; // Replace with the actual folder name
  // Assuming there's only one folder with this name

  // var folder = DriveApp.getFoldersByName(folderName).next(); 
  var folder = getDriveFolderByIdForPDF();  

  // Search for the file with the specified name within the folder
  var files = folder.getFilesByName(fileName);

  // Check if the file exists
  if (files.hasNext()) {
    var file = files.next();
    // Delete the file
    file.setTrashed(true);
    Logger.log("File '" + fileName + "' deleted from folder '" + folderName + "'");
  } else {
    Logger.log("File '" + fileName + "' not found in folder '" + folderName + "'");
  }
}


function getCurrentDateAsmmyy() {
  var currentDate = new Date();
  
  // Get the month and year
  var month = ('0' + (currentDate.getMonth() + 1)).slice(-2); // Adding 1 because months are zero-based
  var year = currentDate.getFullYear().toString().slice(-2);
  
  // Combine them in the "mm.yy" format
  var formattedDate = month + '.' + year;
  
  Logger.log("Current Date: " + formattedDate);
  
  // Optionally, return the formatted date
  return formattedDate;
}


function getOtherMonthDateAsmmyy() {

  var currentDate = new Date();    
  var sheet = SpreadsheetApp.getActiveSheet();
  var currentCell = sheet.getSelection().getCurrentCell();
  var dayno = currentCell.getValue();
  var month = dayno.toString().padStart(2, '0');
  var year = currentDate.getFullYear().toString().slice(-2);  
  var formattedDate = month + '.' + year;
  Logger.log("Current Date: " + formattedDate);
  
  return formattedDate;
}


function createblobpdf(sheetName, pdfName) {
  var sourceSpreadsheet = SpreadsheetApp.getActive();
  var sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
  var url = 'https://docs.google.com/spreadsheets/d/' + sourceSpreadsheet.getId() + '/export?exportFormat=pdf&format=pdf' // export as pdf / csv / xls / xlsx
    +    '&size=A4' // paper size legal / letter / A4
    +    '&portrait=true' // orientation, false for landscape
    +    '&fitw=true' // fit to page width, false for actual size
    +    '&top_margin=0.1'
    +    '&bottom_margin=0.1'
    +    '&left_margin=0.1'
    +    '&right_margin=0.1'
    +    '&printnotes=false'    
    +    '&sheetnames=true&printtitle=false' // hide optional headers and footers
    +    '&pagenum=UNDEFINED&gridlines=true' // hide page numbers and gridlines
    +    '&fzr=false' // do not repeat row headers (frozen rows) on each page
    +    '&horizontal_alignment=CENTER' //LEFT/CENTER/RIGHT
    +    '&vertical_alignment=TOP' //TOP/MIDDLE/BOTTOM
    +    '&gid=' + sourceSheet.getSheetId(); // the sheet's Id
  
  // Show notes
  /*
    +    '&include_notes=false' 
    +    '&includenotes=false'       
    +    '&shownotes=false'    
    +    '&show_notes=false'   
  */

  var token = ScriptApp.getOAuthToken();
  // request export url
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  var theBlob = response.getBlob().setName(pdfName);
  return theBlob;
};


function getDriveFolderByIdForPDF() {
  // Sunbathu has shared this folder will them publicly with Edit permission
  // https://drive.google.com/drive/folders/1pCVXY-HIfDOlZHWWubEvur5eG92JhRhV?usp=sharing

  var sharedFolder = DriveApp.getFolderById('1pCVXY-HIfDOlZHWWubEvur5eG92JhRhV');
  return sharedFolder;
}

function getDriveFolderByName(foldname) {
  // Specify the folder name you want to find
  var folderName = foldname;
  // https://drive.google.com/drive/folders/1pCVXY-HIfDOlZHWWubEvur5eG92JhRhV?usp=sharing
  // var sharedFolder = DriveApp.getFolderById('1pCVXY-HIfDOlZHWWubEvur5eG92JhRhV');
  // return sharedFolder;
 
  // Get the root folder of your Google Drive
  var rootFolder = DriveApp.getRootFolder();  

  // Search for the folder with the specified name
  var folders = rootFolder.getFoldersByName(folderName);

  // Check if the folder exists
  if (folders.hasNext()) {
    var folder = folders.next();
    Logger.log("Found folder: " + folder.getName());
    return folder;
    // Now, you can work with the 'folder' object
    // For example, you can list its files, create new files, etc.

  } else {
    Logger.log("Folder not found: " + folderName);
    return null;
  }
}



function Del_createPdfOfActiveSheet() {
  // Get the active sheet.
  var sheet = SpreadsheetApp.getActiveSheet();

  // Create the export URL.
  var url = "https://docs.google.com/spreadsheets/d/" + sheet.getSpreadsheetId() + "/export?";
  var exportOptions = "exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true";

  // Fetch the PDF.
  var response = UrlFetchApp.fetch(url + exportOptions, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  // Save the PDF to the Drive.
  var blob = response.getBlob();
  var fileName = sheet.getName() + ".pdf";
  blob.save(fileName);
}

function Del_createPdf() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the active sheet
  var sheet = spreadsheet.getActiveSheet();
  
  // Set the PDF export options
  var pdfOptions = {
    pageSize: 'A4',
    fitw: true,  // Fit to page width
    sheetId: sheet.getSheetId()
  };
  
  // Generate the PDF file
  var blob = spreadsheet.getBlob().getAs('application/pdf', pdfOptions);
  
  // Create a new file in Google Drive and save the PDF
  var folder = DriveApp.createFolder('PDFs'); // Change 'PDFs' to your desired folder name
  folder.createFile(blob);
  
  // Display a link to the created PDF
  var pdfFile = folder.getFiles()[0];
  Logger.log('PDF file created: ' + pdfFile.getName());
}




function TestRegex(){

    var data = "first 10.10.23 Yes\nsecond 10.10.23 No\n\n10.10.23 Skip paras\n10.10.23 Need this 10.10.23 wanted";
    var regexPattern = `\\b10\\.10\\.23\\b\\ *(?!.*(No|Created)\\b)\\w+/g`;
    var regexPattern = `\\b10\\.10\\.23\\b\\ *\\w+/g`;
    var regexPattern = `\\b10\\.10\\.23\\b\\ *\\b/g`;
    var regexPattern = `\\b10\\.10\\.23\\b\\ */g`;
    var regexPattern = `\\b10\\.10\\.23\\b\\ *`;
    var regexPattern = `\\bhasan`;
    var regexPattern = `\\b10\\.10\\.23\\b\\ *\\b\\w+`;
    var regexPattern = `\\b10\\.10\\.23\\b\\ *(?!.*(No|Created)\\b)\\w+`;
    var regexPattern = `\\b10\\.10\\.23`;
    var regexPattern = `10\\.10\\.23\\b\\ *\\b\\w+`;
    var regexPattern = `10\\.10\\.23\\b\\ *(?!.*(No|Created)\\b)\\w+`;
    var regexPattern = `10\\.10\\.23\\b\\w+`;
    var regexPattern = `\\b10\\.10\\.23\\s+(\\S+)`;
    var regexPattern = `\\b10\\.10\\.23\\b`;
    var regexPattern = `10\\.10\\.23`;
    var regexPattern = `\\b10\\.10\\.23\\b\\ *(?!.*(No|Skip)\\b)\\w+/g`;
    var regexPattern = '\\b10';
    var regexp = /\+?(\d{1,2}?)(?: *\()?(\d{3})(?:\) *)?(\d{3})-?(\d{2})-?(\d{2}\b)/
    var regexPattern = /\b10\.10\.23\b/

    var regexPattern = /\b10\.10\.23\b\ *(?!.*(No|Created)\b)\w+/
    var regexPattern = /\b10\.10\.23\s+(\S+)/
    var regexPattern = /\b10\.10\.23\ +\w+/
    var regexPattern = /(\b10\.10\.23\ +)(\w+)/

    var regExp = new RegExp(regexPattern, "gi"); // "gi" global & case-insensitive
    // var result = regExp.exec(data)[0];

    // If you want to get results as multiple matches, use () for every single match
    while (match = regExp.exec(data)) {      
      Logger.log(`match[1]: ${match[0]}`);     
      Logger.log(`match[1]: ${match[1]}`);   
      Logger.log(`match[2]: ${match[2]}`);
      // break;
      var YesNoOther = match[2];
    } 

    // Logger.log(`result: ${result}`);
    // Logger.log(`result.Length: ${result.length}`);

}



