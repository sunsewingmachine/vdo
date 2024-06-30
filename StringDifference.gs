// No function is used from this file

function Del_ChangeSingleCharInCell() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C3').activate();
  spreadsheet.getCurrentCell().setRichTextValue(
   SpreadsheetApp.newRichTextValue()
  .setText('test')
  .setTextStyle(1, 2, SpreadsheetApp.newTextStyle()
  .setForegroundColor('#ff0000')
  .build())
  .build());
  spreadsheet.getRange('D3').activate();
};


function stringComparison(s1, s2) {
  // lets test both variables are the same object type if not throw an error
  if (Object.prototype.toString.call(s1) !== Object.prototype.toString.call(s2)){
    throw("Both values need to be an array of cells or individual cells")
  }
  // if we are looking at two arrays of cells make sure the sizes match and only one column wide
  if( Object.prototype.toString.call(s1) === '[object Array]' ) {
    if (s1.length != s2.length || s1[0].length > 1 || s2[0].length > 1){
      throw("Arrays of cells need to be same size and 1 column wide");
    }
    // since we are working with an array intialise the return
    var out = [];
    for (r in s1){ // loop over the rows and find differences using diff sub function
      out.push([diff(s1[r][0], s2[r][0])]);
    }
    return out; // return response
  } else { // we are working with two cells so return diff
    return diff(s1, s2)
  }
}
 
function diff (s1, s2){
  var out = "[ ";
  var notid = false;
  // loop to match each character
  for (var n = 0; n < s1.length; n++){
    if (s1.charAt(n) == s2.charAt(n)){
      out += "–";
    } else {
      out += s2.charAt(n);
      notid = true;
    }
out += " ";
  }
  out += " ]"
  return (notid) ? out :  "[ id. ]"; // if notid(entical) return output or [id.]
}