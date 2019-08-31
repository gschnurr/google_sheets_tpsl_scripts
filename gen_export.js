/** @OnlyCurrentDoc */


//GEN EXPORT

function gen_export_ins() {
  var htmlOutput = HtmlService
    .createHtmlOutput ('<p>1) From the Drop Down Select the Columns you want to export. </p>' +
    '<p> 2) Navigate to the Simple Exports menu item and select Generic Export from the drop down menu. This Macro will produce an output file located in your <a href="https://docs.google.com/spreadsheets/">google sheets drive.</a> </p>' +
    '<p> 3) Open the Spreadsheet titled Generic Export with Todays Date. Within 30 seconds of opening you will be prompted for a macro to run called clean_export. If you are not prompted please select the Export Macros menu button (next to Help) and click Clean Export.</p>'
  )
    .setTitle('Simple Export Instructions');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function gen_export_wrapper() {
  var currentUser = Session.getActiveUser().getEmail();
  gen_export();
  MailApp.sendEmail('gibson.schnurr@izettle.com',
            'General Export Created',
            'The general export macro was run. The running user was ' + currentUser + '.');
  SpreadsheetApp.getUi().alert('Export Created, Please check your google sheet files for the Generic Export with Todays Date');
}

function gen_export() {

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


  //create extract sheet and name
  spreadsheet.insertSheet(2);
  spreadsheet.getActiveSheet().setName('Generic Export');
  var genX = ss.getSheetByName('Generic Export');

  //paste value of specified range
  genX.getRange('A1').activate();
  tpslAllCells.copyTo(genX.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  //deleting TPSL category title row as it is not needed
  genX.getRange('1:1').activate();
  genX.deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());

  //Copying format of title row from tpsl to ppe
  genX.getRange('A1').activate();
  tpsl.getRange(2, 1, 1, tpslLc).copyTo(genX.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  //freezing title row of ppe
  genX.setFrozenRows(1);

  //deleting un-needed columns
  var arrAdj = 1;

  //This array contains all of the columns that you want to keep in the extract
  //If you would like a new column added please add the column header exactly as it is into the array

  function flatten_input_cols_arr() {
    var expGenFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var expGen = ss.getSheetByName('Export Generator');
    var expGenLr = expGen.getLastRow();
    var findTitleColumnArr = expGen.getRange(3, 1, expGenLr, 1).getValues();

    for (row = 0; row < findTitleColumnArr.length; row++) {
      for (column = 0; column < findTitleColumnArr[row].length; column++) {
        expGenFlat.push(findTitleColumnArr[row][column]);
      }
    }
    return expGenFlat
  }

  var expGenOned = flatten_input_cols_arr();

  function flatten_genx_cols_arr() {
    var genXFlat = [];
    var row, column;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var genX = ss.getSheetByName('Generic Export');
    var genXLc = genX.getLastColumn();
    var findTitleColumnArr = genX.getRange(1, 1, 1, genXLc).getValues();

    for (row = 0; row < findTitleColumnArr.length; row++) {
      for (column = 0; column < findTitleColumnArr[row].length; column++) {
        genXFlat.push(findTitleColumnArr[row][column]);
      }
    }
    return genXFlat
  }

  var genXOned = flatten_genx_cols_arr();

  for (var d = 0; d < genXOned.length; d++) {
    if (expGenOned.indexOf(genXOned[d]) == -1) {
      var colPos = d + arrAdj;
      genX.deleteColumn(colPos);
      --arrAdj;
    }
  }

  //!!Note: The newly created sheet is not part of the array because the array is technically created before the ppe sheet was created!!
  //looping through sheets array to protect and hide
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.copy('Generic Export ' + date);

  spreadsheet.deleteSheet(genX);
};

function clean_export_wrapper() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetName = spreadsheet.getName();
  var currentUser = Session.getActiveUser().getEmail();

  if (spreadsheetName == 'TPSL 2.0') {
    SpreadsheetApp.getUi().alert('Error: this macro cannot be run in the master TPSL document, please run this only in your generated export.');
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('WARNING: Google Sheets is set to run the clean_export macro. Confirm you are not in the Master TPSL document, this macro should only be run in generated exports. Are sure you want to continue with this Macro?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {
    return;
  }
  clean_export();
  MailApp.sendEmail('gibson.schnurr@izettle.com',
            'General Export Cleaned',
            'The clean export macro was run on ' + spreadsheetName + '. The running user was ' + currentUser + '.');
}

function clean_export() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var spreadsheet = SpreadsheetApp.getActive();
  for (var i=0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() != 'Generic Export') {
      sheets[i].activate();
      spreadsheet.deleteActiveSheet();
    }
  }
}
