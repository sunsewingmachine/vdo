
function refresh(){
   if (sheetName == "Base") return;
  if (sheetName == "Base1") return;
  if (sheetName == "Base2") return;
  if (sheetName == "Base3") return;
  if (sheetName == "Base4") return;
  if (sheetName == "Base5") return;
  if (sheetName == "Base6") return;
  if (sheetName == "Base7") return;
  if (sheetName == "Base8") return;
  if (sheetName == "Bank") return;
  if (sheetName == "VDoRep") return;
  
  // Utilities.sleep(2000);
  var sheet = SpreadsheetApp.getActiveSheet();
  var totalCols = sheet.getMaxColumns();
  Logger.log("Refreshing, totalCols: " + totalCols.toString());

  for (let k = 2; k < 17; k++) 
  {
    var range = sheet.getRange(k, 3);
    var t1 = range.getValues();
    var jobTime = new Date(t1);

    var d = new Date();
    var timeNow = d.getTime();
        
    var rangeWork = sheet.getRange(k, 2);
    var workName = rangeWork.getValue();     
    var workStatus = sheet.getRange(k, 7).getValue().toString();

    var rng = sheet.getRange(k, 1, 1, totalCols);
    var rowAry = rng.getValues();
    
  // var rng = sht.getRange(rownumber, 1, 1, numberofcolums)
  // var rangeArray = rng.getValues();

    Logger.log("Wname: " + rowAry[0][1]);


    continue;    
    // var workDone = rangeDone.getValue().toString();      
    var rangeRowx = sheet.getRange(k,1,1,totalCols);   //rowfull = '4:4';

    if (jobTime > timeNow)
    {
      Logger.log(workName + ", Done: " + t1.toString());
      // rangeRowx.setBackground("#c3d7e3");   
      sheet.hideRows(k);   
    }
    else
    {
      Logger.log(workName + ", NotDone: " + t1.toString());
      // rangeRowx.setBackground("white");   
      var rangeToUnHide = sheet.getRange(k, 1);   
      sheet.unhideRow(rangeToUnHide);   
    }
  }

}

function myFunction() {
  
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheets()[0];
  // Hides the first row

return;
  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('10:10').activate();
  spreadsheet.getActiveRangeList().setBackground('#d9ead3');

    return;
    var range = SpreadsheetApp.getCurrentCell();
    var currentRow = range.getRow();
    var currentCol = range.getColumn();
    var oldValue = range.getValue();
    var sheet = SpreadsheetApp.getActiveSheet();
    var selectedRows = range.getNumRows(); // Returns the number of rows in this range.
    var selectedCols = range.getNumColumns(); // Returns the number of rows in this range.
    
  /*
  var selection = SpreadsheetApp.getSelection();
  // Current cell: C1
  var currentCell = selection.getCurrentCell();
  // Active Range: C1:D4
  var activeRange = selection.getActiveRange();
  */

  // return;  
  // var data = sheet.getDataRange();  
  // data.activate;
  
  
  // Set background to red if a single empty cell is selected.
  var range = e.range;
  if(range.getNumRows() === 1 
      && range.getNumColumns() === 1 
      && range.getCell(1, 1).getValue() === "") {
    range.setBackground("red");
  }
}


function del(){
/*
       return;
        Logger.log("onEdit Started");
        // var curCell = SpreadsheetApp.getCurrentCell();
        // Logger.log("curCell: " + curCell);
        Logger.log("Range.getColumn(): " + range.getColumn());
        range.setValue("Done");
          
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        // Returns the current highlighted cell in the one of the active ranges.
        var currentCell = sheet.getCurrentCell();
        Logger.log("curCell: " + currentCell.getA1Notation());

        // Logger.log("Range.getColumn()" + range.getColumn());
        range.setValue("Done");

        return;
        Logger.log("ahello farook");  

        if (range.getValue = "Done")
        {
          // range.setValue("Weekly");
           range.setValue("Not ds done");
        }
    }

  // return;
  Logger.log(JSON.stringify(e));
  // Logger.log(e);  //take this value to create dummye
  // return;    

  // Set a comment on the edited cell to indicate when it was changed.
  var range = e.range;
  range.setValue("jone1");
  Logger.log("ahello farook");  
  e.source.toast("hi");  
  range.setValue("jone24");
  return;

  // Set a comment on the edited cell to indicate when it was changed.
  var range = e.range;
  // range.setNote('Last modified: ' + new Date());
  // range.setBackground("green");


  return;

    Logger.log("onEdit Started");
    var range = e.range;
    var currentRow = range.getRow();
    var currentCol = range.getColumn();
    var oldValue = range.getValue();
    var sheet = SpreadsheetApp.getActiveSheet();
    var selectedRows = range.getNumRows(); // Returns the number of rows in this range.
    var selectedCols = range.getNumColumns(); // Returns the number of rows in this range.

    if(selectedRows === 1 && selectedCols === 1 )  { 
      // do      
    } else return;

    Logger.log("range.getNumRows() === 1 && range.getNumColumns() === 1");
    Logger.log("Range.getColumn(): " + range.getColumn());

    if (currentCol === 7)
    {               
      Logger.log("Range.getValue(): " + oldValue);

      if (oldValue == "Done")
      {
        // range.setValue("Not done");                        
        sheet.hideRows(currentRow, 1);
      }
    }
    
  refresh2();
  return;


function test_onEdit() {
  // Utilities.sleep(5000);

  // https://docs.google.com/spreadsheets/d/1jQ-6typqFmEAk6wtlJKdxlw-yWD7cZo-OuAUsqSrzKI/edit#gid=1857499633

 
  // var ss = SpreadsheetApp.openById("1jQ-6typqFmEAk6wtlJKdxlw-yWD7cZo-OuAUsqSrzKI");
  // var sheet = ss.getSheetByName("Base");  // Need to get the sheet, not  just the whole work book.
  // var name1 = sheet.getSheetName();
  // Logger.log("SheetName1: " + name1);
  // sheet.activate;


  var ac = SpreadsheetApp.getActiveSpreadsheet();
  var acSheet = ac.getActiveSheet();
  var name2 = ac.getActiveSheet().getName();    
  Logger.log("Active Sheet Name: " + name2);

  // The code below sets range B15:B15 in the first sheet as the active range.
  var range = acSheet.getRange('B15:B15');
  SpreadsheetApp.setActiveRange(range);  
  var name3 = acSheet.getActiveCell().getA1Notation();   
  Logger.log("Active Cell/Range: " + name3); 

  onEdit({
    user : Session.getActiveUser().getEmail(),
    source : SpreadsheetApp.getActiveSpreadsheet(),
    range : SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
    value : SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
    authMode : "LIMITED"
  });
}


  
  */



}

