/** @OnlyCurrentDoc */

//I might need to add data validation support

function pp_information_prompt() {
  var htmlOutput = HtmlService
    .createHtmlOutput ('<p>1) Make sure that 1_Business Systems Sheet does not have a filter and that the sheet name has not changed </p>' +
    '<p> 2) Navigate to the PP Exports menu item and select PayPal GDPR Extract from the drop down menu. This will run the PayPal GDPR Macro. </p>' +
    '<p> 3) This will take about 10-30 seconds to run and will produce an output file located in your <a href="https://docs.google.com/spreadsheets/">google sheets drive.</a> </p>' +
    '<p> 4) This extract should be passed to legal for the BOs to review. The BO should make updates directly into the extract in the extract columns and mark the updates column as Y if they make updates and N if they do not make any updates.</p>' +
    '<p> 5) Legal will return the extract to you. Once you recieve the extract copy the PayPal Extract sheet to the TPSL spreadsheet (or your copy of it) </p>' +
    '<p> 6) Navigate to the PP Exports menu item and select Get Updates from the drop down menu. This will create a new sheet in the TPSL called Review updates which will contain the updates from the BOs and the original values for that record. </p>' +
    '<p> 7) Review the Changes. If the changes look appropriate do nothing. A Y must be in the last column of the yellow highlighted row if you want to accept the changes and leave the original line item row last column for that record blank. If you do not accept the changes please place a X in both the yellow highlighted row and its corresponding original row for Updates column (last column)</p>' +
    '<p> 8) Navigate to the PP Exports menu item and select Push Updates from the drop down menu. This will add all of the records to the change log sheet and push the updates into the appropraite records in the TPSL.</p>' +
    '<p> 9) Repeat step one to generate a clean file for legal.</p>'
  )
    .setTitle('PayPal Quarterly Extract Instructions');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function tpsl_pp_extract() {
  //general variables
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  //time and date variables
  var tz = ss.getSpreadsheetTimeZone();
  var date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  //tpsl variables
  var tpsl = ss.getSheetByName('1_Business Systems');
  var tpslLr = tpsl.getLastRow();
  var tpslLc = tpsl.getLastColumn();
  var tpslAllCells = tpsl.getRange(1, 1, tpslLr, tpslLc);

  //finding gdpr column and creating a variable on the gdpr column number
  function flatten_tca() {
    var tcaFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tpsl = ss.getSheetByName('1_Business Systems');
    var tpslLc = tpsl.getLastColumn();
    var findTitleColumnArr =tpsl.getRange(2, 1, 1, tpslLc).getValues();

    for (row = 0; row < findTitleColumnArr.length; row++) {
      for (column = 0; column < findTitleColumnArr[row].length; column++) {
        tcaFlat.push(findTitleColumnArr[row][column]);
      }
    }
    return tcaFlat
  }

  var tcaOned = flatten_tca();

  for (var i = 0; i < tcaOned.length; i++) {
    if (tcaOned[i] === 'GDPR Data (Y,N)') {
      var gdprColPos = i + 1; //add one because arrays start at 0 not 1
    }
    else {
      continue;
    }
  }

  //filter tpsl sheet on gdpr column
  tpslAllCells.activate();
  tpsl.setCurrentCell(spreadsheet.getRange('A1'));
  tpslAllCells.createFilter();
  tpsl.getRange('A1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['', 'N', '#N/A']).build();
  tpsl.getFilter().setColumnFilterCriteria(gdprColPos, criteria);
  tpsl.getRange('A1').activate();

  //create extract sheet and name
  spreadsheet.insertSheet(2);
  spreadsheet.getActiveSheet().setName('PayPal Extract');
  var ppe = ss.getSheetByName('PayPal Extract');

  //paste value of specified range
  ppe.getRange('A1').activate();
  tpslAllCells.copyTo(ppe.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  //deleting TPSL category title row as it is not needed
  ppe.getRange('1:1').activate();
  ppe.deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());

  //Copying format of title row from tpsl to ppe
  ppe.getRange('A1').activate();
  tpsl.getRange(2, 1, 1, tpslLc).copyTo(ppe.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  //freezing title row of ppe
  ppe.setFrozenRows(1);

  //deleting un-needed columns
  var arrAdj = 1;

  //This array contains all of the columns that you want to keep in the extract
  //If you would like a new column added please add the column header exactly as it is into the array

  var ppeColsArr = ['SL-ID', 'Application', 'Supplier (Third Party Vendor)', 'Application Manager',
  'Business System Owner', 'GDPR Data (Y,N)', 'Employee Data', 'End Customer Data', 'Merchant Data',
  'Vendor Category', 'Purpose', 'Data Disclosed', 'Data shared with third party? (Y,N,N/A)',
  'Headquarter location'];

  for (var d = 0; d < tcaOned.length; d++) {
    if (ppeColsArr.indexOf(tcaOned[d]) == -1) {
      var colPos = d + arrAdj;
      ppe.deleteColumn(colPos);
      --arrAdj;
    }
  }

  // this inserts an additional column with the specified header
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  ppe.insertColumnAfter(ppeLc);
  ppe.getRange(1, (ppeLc + 1), 1, 1).setValue('Updates? (Y/N) If yes please make the updates in this sheet');

  //!!Note: The newly created sheet is not part of the array because the array is technically created before the ppe sheet was created!!

  //looping through sheets array to protect and hide
  for (var i=0; i < sheets.length; i++) {
    sheets[i].protect();
    sheets[i].hideSheet();
  }

  //copying entire spreadsheet for it to become the paypal extract in copying users drive
  spreadsheet.copy('PayPal Extract ' + date);

  //looping through sheets array to unprotect nad unhide
  for (var i=0; i < sheets.length; i++) {
    sheets[i].showSheet().activate();
    sheets[i].protect().remove();
  }

  //reseting to tpsl to before macro by deleting extract sheet and removing filter
  spreadsheet.deleteSheet(ppe);
  tpsl.getFilter().remove();
  sheets[0].activate();

  //alert for end of macro
  SpreadsheetApp.getUi().alert('Extract Created, Please check your google sheet files for the PayPal Extract with Todays Date');
};




function get_updates() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  //change the name of the copied sheet
  ss.getSheetByName('Copy of PayPal Extract').setName('PayPal Extract');

  var ppe = ss.getSheetByName('PayPal Extract');

  //ppe variables
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  var ppeAllCells = ppe.getRange(1, 1, ppeLr, ppeLc);

  // find the updates column and create a variable with the integer of the column position
  function flatten_ppe_arr_rus() {
    var parFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ppe = ss.getSheetByName('PayPal Extract');
    var ppeLc = ppe.getLastColumn();
    var findTitleColumnArr =ppe.getRange(1, 1, 1, ppeLc).getValues();

    for (row = 0; row < findTitleColumnArr.length; row++) {
      for (column = 0; column < findTitleColumnArr[row].length; column++) {
        parFlat.push(findTitleColumnArr[row][column]);
      }
    }
    return parFlat
  }

  var parOned = flatten_ppe_arr_rus();

  for (var i = 0; i < parOned.length; i++) {
    if (parOned[i] === 'Updates? (Y/N) If yes please make the updates in this sheet') {
      var updatesColPos = i + 1; //add one because arrays start at 0 not 1
    }
    else {
      continue;
    }
  }

  //filter ppe sheet on updates column
  ppeAllCells.activate();
  ppe.setCurrentCell(spreadsheet.getRange('A1'));
  ppeAllCells.createFilter();
  ppe.getRange('A1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['', 'N']).build();
  ppe.getFilter().setColumnFilterCriteria(updatesColPos, criteria);
  ppe.getRange('A1').activate();

  //create extract sheet and name of updates
  spreadsheet.insertSheet(2);
  spreadsheet.getActiveSheet().setName('Review Updates');
  var rus = ss.getSheetByName('Review Updates');

  //paste value of specified range
  rus.getRange('A1').activate();
  ppeAllCells.copyTo(rus.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  //Copying format of title row from tpsl to ppe
  rus.getRange('A1').activate();
  ppe.getRange(1, 1, 1, ppeLc).copyTo(rus.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  //columns and rows to be highlighted
  var rusNumR = rus.getLastRow();
  var rusNumC = rus.getLastColumn();

  //loop through all cells in the range and highlighting them
  for (var i = 2; i <= rusNumR; i++) {
    for (var j = 1; j <= rusNumC; j++) {
      rus.getRange(i, j).setBackground('yellow');
    }
  }
  //get originals

  //tpsl variables and finding the integer for gdpr column
  var tpsl = ss.getSheetByName('1_Business Systems');
  var tpslLr = tpsl.getLastRow();
  var tpslLc = tpsl.getLastColumn();
  var tpslRange = tpsl.getRange(4, 1, tpslLr, 1);
  var tpslArray = tpslRange.getValues();
  var tpslStartRow = tpslRange.getRow();

  function flatten_tca() {
    var tcaFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tpsl = ss.getSheetByName('1_Business Systems');
    var tpslLc = tpsl.getLastColumn();
    var findTitleColumnArr =tpsl.getRange(2, 1, 1, tpslLc).getValues();

    for (row = 0; row < findTitleColumnArr.length; row++) {
      for (column = 0; column < findTitleColumnArr[row].length; column++) {
        tcaFlat.push(findTitleColumnArr[row][column]);
      }
    }
    return tcaFlat
  }

  var tcaOned = flatten_tca();

  for (var i = 0; i < tcaOned.length; i++) {
    if (tcaOned[i] === 'GDPR Data (Y,N)') {
      var tpslGdprBegCol = i + 1; //add one because arrays start at 0 not 1
    }
    else {
      continue;
    }
  }

  var tpslGdprEndCol = 9;

  //creating a one Dimensional array of the IDs in the tpsl sheet
  function flatten_tpsl() {
    var tpslFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tpsl = ss.getSheetByName('1_Business Systems');
    var tpslLr = tpsl.getLastRow();
    var tpslRange = tpsl.getRange(4, 1, tpslLr, 1);
    var tpslArray = tpslRange.getValues();

    for (row = 0; row < tpslArray.length; row++) {
      for (column = 0; column < tpslArray[row].length; column++) {
        tpslFlat.push(tpslArray[row][column]);
      }
    }
    return tpslFlat;
  }

  var tpslOned = flatten_tpsl();

  //rus array and identification
  var rus = ss.getSheetByName('Review Updates');

  function flatten_rustca() {
    var rustcaFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rus = ss.getSheetByName('Review Updates');
    var rusLc = rus.getLastColumn();
    var findTitleColumnArr =rus.getRange(1, 1, 1, rusLc).getValues();

    for (row = 0; row < findTitleColumnArr.length; row++) {
      for (column = 0; column < findTitleColumnArr[row].length; column++) {
        rustcaFlat.push(findTitleColumnArr[row][column]);
      }
    }
    return rustcaFlat
  }

  var rustcaOned = flatten_rustca();

  for (var i = 0; i < rustcaOned.length; i++) {
    if (rustcaOned[i] === 'GDPR Data (Y,N)') {
      var rusGdprBegCol= i + 1; //add one because arrays start at 0 not 1
    }
    else {
      continue;
    }
  }
  var rusGdprEndCol = 9;

  function flatten_rus() {
    var rusFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rus = ss.getSheetByName('Review Updates');
    var rusLr = rus.getLastRow();
    var rusRange = rus.getRange(2, 1, rusLr, 1);
    var rusArray = rusRange.getValues();

    for (row = 0; row < rusArray.length; row++) {
      for (column = 0; column < rusArray[row].length; column++) {
        rusFlat.push(rusArray[row][column]);
      }
    }
    return rusFlat;
  }

  var rusOned = flatten_rus();
  rusOned.pop(); //removes the last item of the array since it is a blank item
  var rusLr1 = rus.getLastRow();
  var rusRange = rus.getRange(2, 1, rusLr1, 1);


  //finding the original information looping and pasting
  //not exactly working
  for (var i = 0; i < rusOned.length; i++) {
    if (tpslOned.indexOf(rusOned[i]) > -1) {
      var rusLr = rus.getLastRow();
      var rusStartRow = rusRange.getRow();
      var rusRow = (i + rusStartRow);
      var tpslIndex = tpslOned.indexOf(rusOned[i]);
      var tpslRow = (tpslIndex + tpslStartRow);
      var tpslCopyRange = tpsl.getRange(tpslRow, tpslGdprBegCol, 1, tpslGdprEndCol);
      var tpslCopyData = tpslCopyRange.getValues();
      var rusConstants = rus.getRange(rusRow, 1, 1, 5).getValues();
      var rusNr = rusLr + 1;

      rus.getRange(rusNr, 1, 1, 5).setValues(rusConstants);
      rus.getRange(rusNr, rusGdprBegCol, 1, rusGdprEndCol).setValues(tpslCopyData);
    }
  }


  //freezing title row of ppe
  rus.setFrozenRows(1);

  //sort to pair ids

  var rusLrSort = rus.getLastRow();
  var rusLcSort = rus.getLastColumn();

  rus.getRange(2, 1, rusLrSort, rusLcSort).sort(1);

  //remove filter on ppe
  ppe.getFilter().remove();
  SpreadsheetApp.getUi().alert('Updates have been pulled');
}

function push_updates() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

//rus sheet vars and functions
  var rus = ss.getSheetByName('Review Updates');
  var rusLr = rus.getLastRow();
  var rusLc = rus.getLastColumn();
  var rusRange = rus.getRange(2, 1, rusLr, 1);
  var rusArray = rusRange.getValues();
  var rusStartRow = rusRange.getRow();


  //find gdpr section start column position
  function flatten_rustca() {
    var rustcaFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rus = ss.getSheetByName('Review Updates');
    var rusLc = rus.getLastColumn();
    var findTitleColumnArr =rus.getRange(1, 1, 1, rusLc).getValues();

    for (row = 0; row < findTitleColumnArr.length; row++) {
      for (column = 0; column < findTitleColumnArr[row].length; column++) {
        rustcaFlat.push(findTitleColumnArr[row][column]);
      }
    }
    return rustcaFlat
  }

  var rustcaOned = flatten_rustca();

  for (var i = 0; i < rustcaOned.length; i++) {
    if (rustcaOned[i] === 'GDPR Data (Y,N)') {
      var rusGdprBegCol= i + 1; //add one because arrays start at 0 not 1
    }
    else {
      continue;
    }
  }
  var rusGdprEndCol = 9;

  // Function that converts 2d array into 1d array
  function flatten_rus() {
    var rusFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rus = ss.getSheetByName('Review Updates');
    var rusLr = rus.getLastRow();
    var rusRange = rus.getRange(2, 1, rusLr, 1);
    var rusArray = rusRange.getValues();

    for (row = 0; row < rusArray.length; row++) {
      for (column = 0; column < rusArray[row].length; column++) {
        rusFlat.push(rusArray[row][column]);
      }
    }
    return rusFlat;
  }

  var rusOned = flatten_rus(); //1 Dimensional array of Column A of the PPE sheet

//tpsl sheet vars and functions
  var tpsl = ss.getSheetByName('1_Business Systems');
  var tpslLr = tpsl.getLastRow();
  var tpslLc = tpsl.getLastColumn();
  var tpslRange = tpsl.getRange(4, 1, tpslLr, 1);
  var tpslArray = tpslRange.getValues();
  var tpslStartRow = tpslRange.getRow();

  //find gdpr section start column
  function flatten_tca() {
    var tcaFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tpsl = ss.getSheetByName('1_Business Systems');
    var tpslLc = tpsl.getLastColumn();
    var findTitleColumnArr =tpsl.getRange(2, 1, 1, tpslLc).getValues();

    for (row = 0; row < findTitleColumnArr.length; row++) {
      for (column = 0; column < findTitleColumnArr[row].length; column++) {
        tcaFlat.push(findTitleColumnArr[row][column]);
      }
    }
    return tcaFlat
  }

  var tcaOned = flatten_tca();

  for (var i = 0; i < tcaOned.length; i++) {
    if (tcaOned[i] === 'GDPR Data (Y,N)') {
      var tpslGdprBegCol = i + 1; //add one because arrays start at 0 not 1
    }
    else {
      continue;
    }
  }

  var tpslGdprEndCol = 9;

// Function that converts 2d array into 1d array
  function flatten_tpsl() {
    var tpslFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tpsl = ss.getSheetByName('1_Business Systems');
    var tpslLr = tpsl.getLastRow();
    var tpslRange = tpsl.getRange(4, 1, tpslLr, 1);
    var tpslArray = tpslRange.getValues();

    for (row = 0; row < tpslArray.length; row++) {
      for (column = 0; column < tpslArray[row].length; column++) {
        tpslFlat.push(tpslArray[row][column]);
      }
    }
    return tpslFlat;
  }

  var tpslOned = flatten_tpsl(); //1 Dimensional array of Column A of the PPE sheet

//x

//create an array of the updates column and find row position and delete the ones that dont equal y
//the updates column needs to be the last column in the sheet
function flatten_rus_updates_col_arr() {
  var rusColUpdatesFlat = [];
  var row, column;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rus = ss.getSheetByName('Review Updates');
  var rusLr2 = rus.getLastRow();
  var rusLcIsUpdatesCol = rus.getLastColumn();
  var rusUpdatesArr = rus.getRange(2, rusLcIsUpdatesCol, rusLr2, 1).getValues();

  for (row = 0; row < rusUpdatesArr.length; row++) {
    for (column = 0; column < rusUpdatesArr[row].length; column++) {
      rusColUpdatesFlat.push(rusUpdatesArr[row][column]);
    }
  }
  return rusColUpdatesFlat
}

var rusUpdatesArrOned = flatten_rus_updates_col_arr();


// BEGINNING OF LOOK UP AND REPLACE
//Use same setup as below to create add the below to an existing 'change log' sheet

  for (var i = 0; i < rusOned.length; i++) {
    if (rusUpdatesArrOned[i] == 'Y') {
      if (tpslOned.indexOf(rusOned[i]) > -1) {
        var rusRow = (i + rusStartRow);
        var rusCopyRange = rus.getRange(rusRow, rusGdprBegCol, 1, rusGdprEndCol);
        var rusCopyData = rusCopyRange.getValues();
        var tpslIndex = tpslOned.indexOf(rusOned[i]);
        var tpslRow = (tpslIndex + tpslStartRow);

        tpsl.getRange(tpslRow, tpslGdprBegCol, 1, tpslGdprEndCol).setValues(rusCopyData);
      }
      else {
        SpreadsheetApp.getUi().alert('could not find ' + rusOned[i] + ' in the TPSL ' + 'make sure that the ID is correct');
      }
    }
    else if (rusUpdatesArrOned[i] == 'X') {
      continue;
    }
    else {
      var gcl = ss.getSheetByName('GDPR Change Log');
      var gclRow = (i +rusStartRow);
      var rusLcGcl = rus.getLastColumn();
      var gclLCopyCol= (rusLcGcl - 1);
      var gclLr = gcl.getLastRow();
      var gclBr = (gclLr + 1);
      var gclCopyData = rus.getRange(gclRow, 1, 1, gclLCopyCol).getValues();
      var gclPasteRange = gcl.getRange(gclBr, 1, 1, gclLCopyCol);

      gclPasteRange.setValues(gclCopyData);
    }
  }

//delete review updates sheet
  var ppe = spreadsheet.getSheetByName('PayPal Extract');
  spreadsheet.deleteSheet(ppe);
  spreadsheet.deleteSheet(rus);
}
