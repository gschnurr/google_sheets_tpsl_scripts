function gen_export_ins() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var htmlOutput = HtmlService
    .createHtmlOutput ('<p>1) From the Drop Down Select the Columns you want to export. </p>' +
    '<p> 2) Navigate to the Simple Exports menu item and select Generic Export from the drop down menu. This Macro will produce an output file located in your <a href="https://docs.google.com/spreadsheets/">google sheets drive.</a> </p>' +
    '<p> 3) Open the Spreadsheet titled Generic Export with Todays Date. Within 30 seconds of opening you will be prompted for a macro to run called clean_export. If you are not prompted please select the Export Macros menu button (next to Help) and click Clean Export.</p>'
  )
    .setTitle('Simple Export Instructions');
  ui.showSidebar(htmlOutput);
}

function gen_export() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
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

  var genXLc = genX.getLastColumn();
  //deleting un-needed columns
  var arrAdj = 1;

  var expGenOned = flatten_arr(expGenColumnArr);

  var genXColumnArr = genX.getRange(1, 1, 1, genXLc).getValues();
  var genXOned = flatten_arr(genXColumnArr);

  for (var d = 0; d < genXOned.length; d++) {
    if (expGenOned.indexOf(genXOned[d]) == -1) {
      var colPos = d + arrAdj;
      genX.deleteColumn(colPos);
      --arrAdj;
    }
  }

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.copy('Generic Export ' + date);

  spreadsheet.deleteSheet(genX);
};

function clean_export() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  for (var i=0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() != 'Generic Export') {
      sheets[i].activate();
      spreadsheet.deleteActiveSheet();
    }
  }
}
