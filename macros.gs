function FormatPage() {
  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().setFrozenRows(1);

  spreadsheet.getActiveRangeList()
  .setBackground('#6aa84f')

  spreadsheet.getRange('D18').activate();
  spreadsheet.getActiveSheet().setColumnWidth(1, 15);
  spreadsheet.getRange('C:C').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getRange('I:I').activate();
  spreadsheet.getActiveRangeList().setBackground('#6aa84f');
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRangeList()
  .setFontSize(16)
  .setFontFamily('Verdana')
  .setFontColor('#ffffff')
  .setHorizontalAlignment('left')
  .setVerticalAlignment('top')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
 
  spreadsheet.getActiveSheet().setRowHeights( 1, 25, 32);



  //Set datavalidation in colStaff.  Now I use this column to move row to another sheet
  //-----------------------------------------------------




  

  spreadsheet.getRange('A1').activate();
};

/*

function Untitledmacro1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B14').activate();
  spreadsheet.getCurrentCell().setValue('helo');
  spreadsheet.getRange('B7').activate();
  spreadsheet.getCurrentCell().setValue('jar');
  spreadsheet.getRange('B8').activate();
};



    continue;    // must use code - don't delete
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
*/ 


function Untitledmacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('10:10').activate();
  spreadsheet.getActiveRangeList().setBackground('#d9ead3');
};

function Untitledmacro2() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() + 2, 1, 1, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRangeList().setBackground('#cfe2f3');
  spreadsheet.getCurrentCell().offset(3, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#fff2cc');
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow() - 1, 1, 1, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRangeList().setBackground('#d9ead3');
  spreadsheet.getCurrentCell().offset(-4, 0).activate();
};

function SetStatusValidation() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(-10, 0, 19, 1).activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Db'), true);
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn(), sheet.getMaxRows(), 1).activate();
  
  spreadsheet.getRange('Base!H2:H20').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(spreadsheet.getRange('Db!$B:$B'), true)
  .build());
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Base'), true);
};


function MacroAddNewRowBelow() {
  // var spreadsheet = SpreadsheetApp.getActive();
  // spreadsheet.getRange('C8').activate();
   NewJob();
};

 
//Reference
//activeSheet.deleteRow(currentRow); 
 

function Add5row() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2').activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 5);
};



function SetF() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B32').activate();
};



function Untitled7866() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(43, -1, 4, 3).activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getCurrentCell().offset(4, 0, 3, 3).activate();
};


