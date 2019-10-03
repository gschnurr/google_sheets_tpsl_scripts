/** @OnlyCurrentDoc */
//Sheet
var spreadsheet = SpreadsheetApp.getActive();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();
var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
var spreadsheetName = spreadsheet.getName();
var sheetName = ss.getSheetName();

function flatten_arr(targetArr) {
 var flatArr = [];
 var row, column;

 for (row = 0; row < targetArr.length; row++) {
   for (column = 0; column < targetArr[row].length; column++) {
     flatArr.push(targetArr[row][column]);
   }
 }
 return flatArr
}


function replace_x() {

var spreadsheet = SpreadsheetApp.getActive();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
var spreadsheetName = spreadsheet.getName();
var sheetName = ss.getSheetName();
var newApps = ss.getSheetByName('NEW APPLICATIONS');
var naLr = newApps.getLastRow();
var naLc = newApps.getLastColumn();
//create an array of all of the col title, use that col title array position to get the column position = array pos + 2
var titleColArr = newApps.getRange(1, 2, 1, naLc).getValues();
var tcaOned = flatten_arr(titleColArr);

 for (var r = 0; r < tcaOned.length; r++) {
   var colPos = r + 2;
   var dataArr = newApps.getRange(3, colPos, naLr, 1).getValues();
   var daOned = flatten_arr(dataArr);

   for (var d = 0; d < daOned.length; d++) {
     var rowPos = d + 3;
     if (daOned[d] == 'X' || daOned[d] == 'x') {
       var replaceCell = newApps.getRange(rowPos, colPos, 1, 1);
       replaceCell.setValue(tcaOned[r]);
     }
     else {
       continue;
     }
   }
 }
}
