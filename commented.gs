function myFunction() {
  
}





/*
  try 
  {
    // code to catch error
  } 
  catch (error) 
  {
    msg = "row=" + row + ", colNoStatus="+ colNoStatus + ", status="
    SpreadsheetApp.getUi().alert(msg + ",  " +error);
    return;
  }
  

function onChangeInstallableTriggers(e)
{    
  var email = Session.getActiveUser().getEmail();
  Logger.log("From onChangeInstallableTriggers, Session.getActiveUser().getEmail(): " + email);  

  var user = e.user;
  Logger.log("From onChangeInstallableTriggers, e.user: " + user );  

  if (user == null || user == "sunbathu@gmail.com")
    return;

  var changeType = e.changeType;

  if(changeType == "REMOVE_ROW" || changeType == "REMOVE_COLUMN")
  {
    var txt1 = 'Warning:\n\nஎதையும் E2 விடம் கேட்காமல் Delete செய்ய வேண்டாம்.\n\nஉடனே Undo செய்யவும்.\n\n' + 
    "Deleted by: You" ;
    var txt2 = 'தீர்வு:\n\nDelete செய்யவதற்கு பதிலாக Status-ஜ Disabled என்று மாற்றவும்.\n';
    SpreadsheetApp.getUi().alert(txt1 + txt2);
    // SpreadsheetApp.getUi().alert();    
  }

  Logger.log("From onChangeInstallableTriggers, e.user: " + user + ", " + "e.changeType: " + changeType);  
}
*/



/*

function Post1day() 
{  
  PostponeDate(1);
};

function Post3days() 
{  
  PostponeDate(3);
};

function Post1week() 
{  
  PostponeDate(7);
};

function PostponeDate(days) 
{  
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var ac = sheet.getActiveCell();  
  var row = ac.getRow();  
  var job = sheet.getRange(row, colNoWork).getValue();
  if (job.toString().length < 1) return;

  if (days < 7)
  {
    var ds = (days > 1 )? " days" : " day"
    sheet.getRange(row, colNoStatus).setValue("Postpone " + days.toString() + ds);    
  }
  else if (days < 22)
  {
    var ws = (days > 7 )? " weeks" : " week"
    sheet.getRange(row, colNoStatus).setValue("Postpone " + (days/7).toString() + ws);    
  }

  // Utilities.sleep(5000);  
  refresh2(row, WorkUpdate);

};


function moveDrawing(){
  const obj = {
                "10": {name: " Drawing1 ", moveTo: [1,1]}
              };
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Farook");

  sheet.getDrawings()
    .forEach(d => {
      const arow = d.getContainerInfo().getAnchorRow();
      
      if (arow in obj) {
        d.setPosition(...obj[arow].moveTo, 10, 100);
      }
    Logger.log(arow);
    }  
  )

   // 3. Move drawings.
  sheet.getDrawings().forEach(d => {
    const arow = d.getContainerInfo().getAnchorRow();
    if (arow in obj) {
      d.setPosition(...obj[arow].moveTo, 0, 0);
    }
  })
  
}
*/

// from onEdit() of code.gs
    /* bcoz i have written this in staff column 
          //Show all rows including hidden. Useful in mobile view. Bcoz mobiles view do not show menus
          if (pv == "all") {        
            activeSheet.deleteRow(currentRow); 
            showAllWorks();
            var spreadsheet = SpreadsheetApp.getActive();
            spreadsheet.getRange('B2').activate();
            return;
          }

          //Hide done works by typing hide in last empty cell. Bcoz mobiles view do not show menus
          if (pv == "hide") {        
            activeSheet.deleteRow(currentRow); 
            HideDoneWorks();    
            var spreadsheet = SpreadsheetApp.getActive();
            spreadsheet.getRange('B2').activate();
            return;
          }
    */