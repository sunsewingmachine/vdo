/*
Word todo by E3:
1) Add 'duplicate' in Col.A for duplicate entries (one with Id, one without id)
2) 

*/

var DoOnlyOnce_IdCreation_InStaffPage_AsPer_BasePage = false;


function DeleteDuplicates(){
  // NOT USED NOW
  const ColToFindDuplicate = 1; // 1 is Column B
  var SourceSheet = SpreadsheetApp.getActiveSheet();
  var SheetName = SourceSheet.getSheetName();
  LoggerLog("SheetName: " + SheetName);     
  if(SheetName != "Bathu") return;

  var dataSource = SourceSheet.getRange("A:F").getValues();
  // LoggerLog("dataSource: " + dataSource);     
  
  for (var sourceRow=1; sourceRow < dataSource.length; sourceRow++) {
    
    var workName = dataSource[sourceRow][ColToFindDuplicate];
    LoggerLog("sourceRow: " + sourceRow + ", workName: " + workName.substring(0, 20));
  }
}


function DoOnceIdCreation(){ 
  // Run this function manually
  // ========================Only Once Per Staff=========================
  // For creating ID of Base page in Staff page for 1st time only!!!!
  // This if block searches for all assigned works by workName in staff page,
  // if found, it puts base page's id in staff page.

  DoOnlyOnce_IdCreation_InStaffPage_AsPer_BasePage = true;
  CheckIfAnyWorkMissing();
}

function GetOrCreateSheet(sheetName){
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

  if (sheet == null){
    sheet = SpreadsheetApp.getActive().insertSheet();
    sheet.setName(sheetName);
  }

  return sheet;
}



function CheckIfAnyWorkMissing()
{
  // Base page ல் B5 என்று குறிப்பிடபட்டுள்ள வேலைகள் B5 page ல் உள்ளதா என்பதை பார்க்கும்.
  // அந்த வேலை இருந்தால் base page row number ஐ அந்த ஸ்டாப் page ல் முதல் காலத்தில் போடும்.
  // அந்த வேலை இல்லாவிட்டால் அந்த ஸ்டாப் page ல் புது ரோ உருவாக்கி அந்த வேலையை போடும். 
  // மேலும் Missing Added என முதல் காலத்தில் போடும்.
  // B5 வேலைகள் அனைத்தும் சரியாக உள்ளதா என செக் செய்யவேண்டுமென்றால் 
  // Base page ல் B5 என்ற காலத்தில் வைத்து compareBaseWorks என்ற மெனுவை கிளிக் செய்க.

  // var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var SourceSheet = SpreadsheetApp.getActiveSheet();
  var SourceSheetName = SourceSheet.getName();
  // var totalRows = SourceSheet.getMaxRows();
  var row = SourceSheet.getActiveRange().getRow();
  var col = SourceSheet.getActiveRange().getColumn();  
  var ui = SpreadsheetApp.getUi();


  if (col <= colNoReport || SourceSheetName != "Base")  {
    showAlert('Sheet must Base and column must be after col.Report.');
    showAlert("Condition1: Base sheet-ல் இருந்து இந்த function run செய்யவும்." + "\n" + "Condition 2: Column: J க்கு பிறகு இருக்கவேண்டும்." + "\n" + "அதாவது யாருக்கு செக் செய்யணுமோ அந்த column-ல் mouse இருக்கனும்");
     var msg = ui.alert("Condition 1: Base sheet-ல் இருந்து இந்த function run செய்ய வேண்டும்." + "\n" + "Condition 2: Column: J க்கு பிறகு curson இருக்கவேண்டும்." + "\n" + "அதாவது யாருக்கு செக் செய்யணுமோ அந்த column-ல் mouse இருக்கனும்", ui.ButtonSet.OK); 
    return;
  }
  
  var DestinationSheetName = SourceSheet.getRange(1,col).getValue();
  var DestinationSheet = SpreadsheetApp.getActive().getSheetByName(DestinationSheetName);
  let data = DestinationSheet.getRange(1, 2, 2000).getValues();

  
  // Get confirmation to proceed
  // ____________________________________________________________________________
  
  var msg = ui.alert("It will compare works form sheet base with sheet: " + DestinationSheetName + "\n" + "After comparing, it will put missing works if any from page base to "  + DestinationSheetName + "\n" + "\n" + "Function run ஆகிமுடித்தவுடன் மறுபடி மெசேஜ் வரும். அதுவரை எதையும் கிளிக் செய்யவேண்டாம்.!" + "\n" + "\n" + "செக் செய்ய விரும்புகிறீர்களா?" , ui.ButtonSet.YES_NO); 

  if (msg == ui.Button.YES) {              
  } 
  else {
    return;
  }  
  

  // PUT NEXT UNIQUEID IN 'A1' CELL OF BASE PAGE
  var uid = GetUniqueIDForNewJob2(SourceSheet);
  // LoggerLog("Uid: " + uid);
  if(uid == 0){
    showAlert("Correct A1 of Base Sheet. It should be an Integer & the lastly used id.");
    return;
  }
  else if(uid == -1){
    // Error in Unique Id of Base sheet so exit function;
    return;
  }      
  PutNextNumberInCell(SourceSheet, 'A1', uid);
  

  
  var sheetMissingWorks = GetOrCreateSheet('MissingWorks');
  sheetMissingWorks.showSheet();

  // var drange = SourceSheet.getDataRange(); 
  // var values = drange.getValues(); // todo

  var values = SourceSheet.getRange("A:D").getValues();
  var data2 = DestinationSheet.getRange("A:A").getValues();
  var data3 = DestinationSheet.getRange("A:D").getValues();
  var dataSource = SourceSheet.getRange("A:D").getValues();
  var colLetter = getColumnLetterFromNumber(col);
  var colFullStaff = colLetter + ':' + colLetter;
  var dataOfStaffColumn = SourceSheet.getRange(colFullStaff).getValues();  
  LoggerLog("colFullStaff: " + colFullStaff ) ;        

  // LoggerLog("dataOfStaffColumn: " + dataOfStaffColumn ) ;      
  // LoggerLog("dataSource: " + dataSource ) ;     
  //  LoggerLog("data3: " + data3 ) ;         
  //  return;

  

  var anyMissedOrMismatch = false;
  var rowWD = 1;  // row in WorksDone sheet
  sheetMissingWorks.getRange(1, 1, sheetMissingWorks.getMaxRows(), sheetMissingWorks.getMaxColumns()).clearContent();

  for (var sourceRow=1; sourceRow<values.length; sourceRow++) {
    var workAssigned = (dataOfStaffColumn[sourceRow]=='s');
    if(workAssigned == false) continue;

    var currentId = values[sourceRow][0];
    if(currentId == ''){            
      var workNameS2 = dataSource[sourceRow][1];
      if(workNameS2 != '') LoggerLog("Empty: WorkId: " + currentId + ", sourceRow: " + (sourceRow+1));
      continue;
    }

    // It will find the job in staff page, using Id from Base page
    var foundRow = data2.findIndex(mId => {return mId[0] == currentId});  
        
    var workMissing = false;
    var workMisMatch = false;

    var workNameD = "";
    var workNameS = dataSource[sourceRow][1]; 


    // ========================Only Once Per Staff Start=========================
    // For creating ID of Base page in Staff page for 1st time only!!!!
    // This if block searches for all assigned works by workName in staff page,
    // if found, it puts base page's id in staff page.

    if(DoOnlyOnce_IdCreation_InStaffPage_AsPer_BasePage == true)
    {
      //It will find the job in staff page, using job name
      var foundRow2 = data3.findIndex(mId => {return mId[1] == workNameS});  

      if(foundRow2 > 0) { // If job found in staff page
        foundRow2 = foundRow2 + 1;
        LoggerLog("foundRow2: " + foundRow2 + ", workNameS: " + workNameS.substring(0, 20));
        DestinationSheet.getRange("A"+ foundRow2).setValue(currentId);
        continue;
      }else{
        continue;
      }
    }
    // ========================Only Once Per Staff End=========================



    // If Base'Id is found in Staff page, foundRow = somevalue
    if(foundRow > 0){ 

      // ============BLINDLY OVERWRITE STAFF'S WORK WITH BASE WORK==============

      var sRow = (sourceRow+1);   
      var sourceRng = "A" + sRow;      
      var sourceRng2 = "B" + sRow;      
      var temp = SourceSheet.getRange(sourceRng).getValue();     
      var temp2 = SourceSheet.getRange(sourceRng2).getValue();     
      var temp2 = temp2.substring(0, 15);     
      LoggerLog("Source: " + sourceRng + ", Id: " + temp + ", D: " + temp2);

      var fRow = (foundRow+1);      
      var destRng = "A" + fRow;
      var destRng2 = "B" + fRow;
      var temp = DestinationSheet.getRange(destRng).getValue();   
      var temp2 = DestinationSheet.getRange(destRng2).getValue();       
      var temp2 = temp2.substring(0, 15);     
      LoggerLog("Destin: " + sourceRng + ", Id: " + temp + ", D: " + temp2);

      var temp = SourceSheet.getRange(sRow, 2, 1, colNoProperTime-1).getValues();      
      DestinationSheet.getRange(fRow, 2, 1, colNoProperTime-1).setValues(temp);      
     
      // =======================CHECK MISMATCH WORKS STARTS=====================================
      // THIS IF BLOCK IS FOR CHECKING workMisMatch WORKS.
      // I.E, WHEN STAFF CHANGES ANY WORK IN HIS SHEET, EXCEPT ONCE WORKS, WHICH HE CAN CHANGE!
      
      if (1 == 2){
        // ================================================================
        // If workId found in destination-staff sheet, get that work details.
        var workNameD = data3[foundRow][1];

        // If work is unchanged in staff page, i.e, BasePage's work == Staff'Page's work
        if(workNameS == workNameD){
          // ok, good, No change in user's sheet
          LoggerLog("Correct: WorkId: " + currentId + ", RowD: " + (foundRow+1) +  ", WorkD: " + 
          workNameD.substring(0, 15) + ", WorkS: " + workNameS.substring(0, 15)) 
          continue; // ok, check next work
        }
        else
        {
          workMisMatch = true;
          LoggerLog("Creating Mismatch Work: WorkId: " + currentId + ", RowD: " + (foundRow+1) + ", WorkS: " + workNameS.substring(0, 12));
        }
      }      
      // =======================CHECK MISMATCH WORKS ENDS=====================================
    }
    else
    { 
      // THIS BLOCK IS EXECUTED WHEN A WORK IS MISSING, AS PER OUR ID
      workMissing = true;
      // LoggerLog("WorkId: " + currentId + ", RowD: " + (foundRow+1) +  ", sourceRow: " +  sourceRow);
      // LoggerLog("workNameS: " + workNameS);
      LoggerLog("Creating Missing Work: WorkId: " + currentId + ", RowS: " + (sourceRow+1) + ", WorkS: " + workNameS.substring(0, 12));
    }

    // Since we overwrite staff's work with base work, workMisMatch is always false
    workMisMatch = false; 
    if(workMissing || workMisMatch)
    {
      anyMissedOrMismatch = true;
      DestinationSheet.activate();
      var rngSourceJob = "A" + (sourceRow+1).toString() + ":" + "F"+ (sourceRow+1).toString();      
      var tempValuesOfSourceJob = SourceSheet.getRange(rngSourceJob).getValues();

      if(workMissing){
        var newDestRow = GetFirstBlankRowInSheet(DestinationSheet);
      }
      else if(workMisMatch) {
        DestinationSheet.insertRowAfter(foundRow+1);
        var newDestRow = foundRow + 2;
      }

      var rngDest = "A"+ newDestRow + ':' + 'F' + newDestRow;
      DestinationSheet.getRange(rngDest).setValues(tempValuesOfSourceJob);
      DestinationSheet.getRange("A"+ newDestRow).activate();      

      if(workMisMatch){
        rowWD++;
        sheetMissingWorks.getRange(rowWD, 1).setValue("Changed, id:" + currentId);
        sheetMissingWorks.getRange(rowWD, 2).setValue((sourceRow+1).toString());      
        sheetMissingWorks.getRange(rowWD, 3).setValue(workNameS);
        sheetMissingWorks.getRange(rowWD, 4).setValue(workNameD);  
      }

      if(workMissing){
        rowWD++;
        sheetMissingWorks.getRange(rowWD, 1).setValue("Missing, id:" + currentId);
        sheetMissingWorks.getRange(rowWD, 2).setValue((sourceRow+1).toString());    
        sheetMissingWorks.getRange(rowWD, 3).setValue(workNameS);        
      }

    }
  }

  SortColumnsBy462(DestinationSheet);
    
  
  // CODE-BLOCK: MISSING-IN-BASE
  // THIS BLOCKS CHECKS FOR WORKS THAT ARE IN STAFF PAGE BUT NOT IN BASE PAGE
  // I.E, STAFF WOULD HAVE CREATED, WITHOUT INFORMING BATHU

  rowWD++;
  for (var row=1; row<data3.length; row++) {
    var workOnce = (data3[row][3] == 'Once');
    if(workOnce == true) continue;
    var UidOfThisWork = data3[row][0];

    if(UidOfThisWork == ''){            
      // ALERT: STAFF PAGE HAS A 'NON-ONCE' WORK, WITHOUT UID!
      var workNameS3 = data3[row][1];
      if(workNameS3 != '') {
        LoggerLog("Missing in Base: Work: " + workNameS3 + ", sourceRow: " + (row+1));
        rowWD++;
        sheetMissingWorks.getRange(rowWD, 1).setValue("New work, not in Base:");
        sheetMissingWorks.getRange(rowWD, 2).setValue(UidOfThisWork.toString());   

        sheetMissingWorks.getRange(rowWD, 4).setValue(workNameS3);  
        sheetMissingWorks.getRange(rowWD, 5).setValue(data3[row][3]);  

        // 'https://docs.google.com/spreadsheets/d/1jQ-6typqFmEAk6wtlJKdxlw-yWD7cZo-OuAUsqSrzKI/edit#gid=2070620945&range=B75'

        var t1 = "https://docs.google.com/spreadsheets/d/1jQ-6typqFmEAk6wtlJKdxlw-yWD7cZo-OuAUsqSrzKI/edit#gid=";
        var t2 = DestinationSheet.getSheetId().toString();   // "2070620945"
        var t3 = "&range=B";
        var t4 = (row+1).toString();
        
        // var linkToCell = "=" + '"' + t1 + t2 + t3 + t4  + '"';     
        // sheetMissingWorks.getRange(rowWD, 3).setValue(linkToCell);

        var linkToCell2 = '"' + t1 + t2 + t3 + t4  + '"';     
        var value = '=HYPERLINK(' + linkToCell2 + ', "Goto that Row")';
        sheetMissingWorks.getRange(rowWD, 3).setFormula(value);
      }
      continue;
    }
  }
  // CODE-BLOCK: MISSING-IN-BASE  OVER
  
  

  if(anyMissedOrMismatch) Del_FormatCellWrapHeightAuto(sheetMissingWorks, DestinationSheet);

  showToast('Done. All works checked with sheet : ' + DestinationSheetName);
  showAlert('Done. All works checked with sheet : ' + DestinationSheetName);
  var msg = ui.alert("Base-ல் உள்ள வேலைகள் சரிபார்ப்பது முடிந்தது. இனி நீங்கள் வேலை செய்யலாம்.!", ui.ButtonSet.OK); 
  return;




  /*
  // B5 என்றால் அந்த காலத்தில் எந்த ரோவில் B5 என்று அசைன் செய்துள்ளோமோ 
  // அந்த வேலைகள் B5 page ல் உள்ளதா? என பார்த்து நம்பர் போடும் 
  // அல்லது வேலையையே காப்பி செய்து போடும்.
  // ____________________________________________________________________________
  
  for (let rw = 2; rw <= totalRows; rw++)
  //  for (let rw = 2; rw <= 10; rw++)
  {
    var valueInColOfThatStaff = SourceSheet.getRange(rw,col).getValue();

    if (valueInColOfThatStaff == DestinationSheetName || valueInColOfThatStaff == "S" || valueInColOfThatStaff == "s")  //Check existing of job (matching of job), only if this job has been assigned to that staff. I Put "S" or that staffpage name in that row.
    {
      //குறிப்பிட்ட வேலை உள்ளது. எனவே அந்த ரோ நம்பரை மட்டும் முதல் காலத்தில் போடு.
      // ____________________________________________________________________________
      var curJob = SourceSheet.getRange(rw,colNoWork).getValue();
      let foundRow = data.findIndex(users => {return users[0] == curJob});                      //It will find the job in staff page.
      if (foundRow > 0){
        SpreadsheetApp.getActive().getSheetByName(DestinationSheetName).getRange(foundRow+1,1).setValue(rw)  //It will put row number of the work in base page to column 1 of staff page
      }
      else{
      //குறிப்பிட்ட வேலை இல்லாததால் அதை காப்பிசெய்து போடு        
      // ____________________________________________________________________________
      // var database = SpreadsheetApp.openById("XXX");
      var source = ss.getSheetByName('Base');
      var dataToCopy = source.getRange('A' + rw + ':' +  'H' + rw);
      // var copyToSheet = database.getSheetByName(DestinationSheetName);
      var sourceValues = dataToCopy.getValues();
      var lastRow = SpreadsheetApp.getActive().getSheetByName(DestinationSheetName).getLastRow();
                    
      SpreadsheetApp.getActive().getSheetByName(DestinationSheetName).getRange(lastRow + 1, 1, sourceValues.length, sourceValues[0].length).setValues(sourceValues);
      SpreadsheetApp.getActive().getSheetByName(DestinationSheetName).getRange(lastRow + 1,1).setValue('RowMissing. So copied from server') 
      //dataToCopy.clear({contentsOnly:true});
      }
    }
    
  }
  */

   

}

function GetFirstBlankRowInSheet(Actsheet){    
  if (typeof Actsheet === 'undefined') { 
    Actsheet = SpreadsheetApp.getActiveSheet();
  }

  Actsheet.insertRowAfter(Actsheet.getMaxRows()); // it ensures a blank row
  var values = Actsheet.getDataRange().getValues();
  var FirstBlankRow = 0;
  for (var FirstBlankRow=0; FirstBlankRow<values.length; FirstBlankRow++) {
    if (!values[FirstBlankRow].join("")) break;
  }
  return (FirstBlankRow +1);
}

/*
function GetUniqueIDForNewJobNotUsed(sheet){  
  var uid = sheet.getRange("A1").getValue();

  if(Number.isInteger(uid)) 
  {
    uid = uid + 1;
    sheet.getRange("A1").setValue(uid);
    return uid;
  }
  else
  {
    var aa = sheet.getRange("A:A").getValues();    
    LoggerLog(aa);
    return 0;
  }
}
*/


function PutNextNumberInCell(sheet, cellAddress, thisNumber){

  if(Number.isInteger(thisNumber)) 
  {
    thisNumber = (thisNumber + 1).toString();
    sheet.getRange(cellAddress).setValue(thisNumber);
    LoggerLog(`PutNextNumberInCell:  nextNumber=${thisNumber}`);
    return thisNumber;
  }
  else
  {
    LoggerLog(`PutNextNumberInCell:  thisNumber=${thisNumber} (not an integet)`);
    return null;
  }
}

function GetUniqueIDForNewJob2(sheet){  
  var prevIds = sheet.getRange("A:B").getValues();
  // LoggerLog(prevIds);
  prevIds.shift(); // theRemovedElement == 1

  var arrayLength = prevIds.length;
  for (var i = 0; i < arrayLength; i++) {
    var eachId = prevIds[i][0];
    if(Number.isInteger(eachId) == false){
      var workName2 = prevIds[i][1];
      if(eachId == '' && workName2 == ''){
        // just an empty row, leave it. no problem.
      }else{
        // LoggerLog("Danger: Id Missing in Base sheet at Row: " + (i+1));
        showAlert("WorkId is missing in Row: " + (i+2));
        return -1;        
      }
      prevIds[i] = 0; // todo: actually there is no need for this, we Must set id!
      continue;
    }
    prevIds[i] = prevIds[i][0];
  }
  var largest = Math.max.apply(0, prevIds);
  var nextId = largest + 1;
  // LoggerLog(prevIds);
  LoggerLog("Next Id: " + nextId);
  return nextId;
}

function toInteger(val){
    val = parseInt(val);    // in case val is not an integer
    return val.toString()
}

function getColumnLetterFromNumber(column){
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


function Del_FormatCellWrapHeightAuto(Actsheet, destSheet) {
  if (typeof Actsheet === 'undefined') { 
    Actsheet = SpreadsheetApp.getActiveSheet();
  }
  Actsheet.activate();
  Actsheet.setColumnWidth(1, 90);
  Actsheet.setColumnWidth(2, 50);
  Actsheet.setColumnWidth(3, 700);
  Actsheet.setColumnWidth(4, 700);  

  Actsheet.getRange("A1").setValue("Status");
  Actsheet.getRange("B1").setValue("Source Row");
  Actsheet.getRange("C1").setValue("Data in Base");
  Actsheet.getRange("D1").setValue("Data in " + destSheet.getName());
  Actsheet.getRange("A1:D1").setBackground("#74b572");
  Actsheet.setFrozenRows(1);
  Actsheet.setRowHeights(1,1,30);

  Actsheet.getRange('C:D').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  Actsheet.getRange('C:D').setVerticalAlignment("top");
  Actsheet.autoResizeRows(2, Actsheet.getMaxRows()-1);  
  Actsheet.getRange("A1").activate();
};

function CheckIfAnyWorkMissingOLD(){

  // The code below gets the values for the range C2:G8
  // in the active spreadsheet.  Note that this is a JavaScript array.
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getLastRow();
  var values = sheet.getRange(2, 1, rows, 1).getValues();

  Logger.log(values[0][0]);
}


function GetNumberIfOrNull(text, returnvalue){
  if(Number.isInteger(text)) 
  {
    return text;
  }
  else
  {
    return returnvalue;
  }
}

