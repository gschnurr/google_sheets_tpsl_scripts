function pp_information_prompt() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

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
  ui.showSidebar(htmlOutput);
}

function tpsl_pp_extract() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var tcaOned = flatten_arr(tpslTitleColumnArr);
  //create extract sheet and name
  spreadsheet.insertSheet(2);
  spreadsheet.getActiveSheet().setName('PayPal Extract');
  var ppe = ss.getSheetByName('PayPal Extract');
  logs_tst('PayPal Extract sheet created.');

  //paste value of specified range this should probably change from a paste function to a set function
  var tpslWholeSheetArr = tpslAllCells.getValues();
  ppe.getRange(1, 1, tpslLr, tpslLc).setValues(tpslWholeSheetArr);
  logs_tst('Values from TPSL copied to the PPE sheet without format.');
  var tpslAllNotes = tpslAllCells.getNotes();
  var tpslAllDvRules = tpslAllCells.getDataValidations();
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  var ppeAllCells = ppe.getRange(1, 1, ppeLr, ppeLc);
  ppeAllCells.setNotes(tpslAllNotes);
  ppeAllCells.setDataValidations(tpslAllDvRules);
  logs_tst('TPSL notes and data validation settings copied to PPE Sheet');

  //deleting TPSL category title row as it is not needed
  ppe.deleteRow(1);
  //Copying format of title row from tpsl to ppe
  ppe.getRange('A1').activate();
  tpsl.getRange(2, 1, 1, tpslLc).copyTo(ppe.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  logs_tst('Format for title row copied from TPSL to PPE Spreadsheet');
  var ppeTcArr = ppe.getRange(1, 1, 1, ppeLc).getValues();
  var ppeTcaOned = flatten_arr(ppeTcArr);
  var ppeGdprColPos = find_col(ppeTcaOned, 'GDPR Data (Y,N)');
  var ppeGdprColFilItArr = ['Y'];
  var gdprYesNoArry = ppe.getRange(2, ppeGdprColPos, ppeLr, 1).getValues();
  var gdprYesNoOned = flatten_arr(gdprYesNoArry);
  logs_tst('The gdpr column is in column ' + ppeGdprColPos + '. The values in that column are ' + gdprYesNoOned);
  filter_rows(gdprYesNoOned, ppeGdprColFilItArr, ppe);
  logs_tst('Row filtering complete');
  //freezing title row of ppe
  ppe.setFrozenRows(1);

  //deleting un-needed columns
  filter_cols(tcaOned, ppeColsArr, ppe);
  logs_tst('Column filtering complete.');

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
  sheets[0].activate();

  var compLogs = Logger.getLog();
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'TPSL GDPR PPE Gen Logs', compLogs);
  //alert for end of macro
  ui.alert('Extract Created, Please check your google sheet files for the PayPal Extract with Todays Date');
};

function get_updates() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  //change the name of the copied sheet
  ss.getSheetByName('Copy of PayPal Extract').setName('PayPal Extract');
  var ppe = ss.getSheetByName('PayPal Extract');
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  var ppeAllCells = ppe.getRange(1, 1, ppeLr, ppeLc);

  var ppeTitleColumnArr = ppe.getRange(1, 1, 1, ppeLc).getValues();
  var parOned = flatten_arr(ppeTitleColumnArr);

  //find the updates column and create a variable with the integer of the column position
  var updatesColPos = find_col(parOned, 'Updates? (Y/N) If yes please make the updates in this sheet');

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

  //Creating a one dim arr of tpsl column headers to find the integer for gdpr column
  var tcaOned = flatten_arr(tpslTitleColumnArr);
  var tpslGdprBegCol = find_col(tcaOned, 'GDPR Data (Y,N)');
  //number of columns in the GDPR column set
  var tpslGdprEndCol = 9;
  //A 1d array of the IDs in the tpsl sheet
  var tpslOned = flatten_arr(tpslArray);
  //rus variables
  var rus = ss.getSheetByName('Review Updates');
  var rusLc = rus.getLastColumn();
  var rusTitleColumnArr = rus.getRange(1, 1, 1, rusLc).getValues();
  var rusLr = rus.getLastRow();
  var rusRange = rus.getRange(2, 1, rusLr, 1);
  var rusArray = rusRange.getValues();
  var rusStartRow = rusRange.getRow();
  // creating a 1d arr of column headers in rus and identification of GDPR column position
  var rustcaOned = flatten_arr(rusTitleColumnArr);
  var rusGdprBegCol = find_col(rustcaOned, 'GDPR Data (Y,N)');
  //number of columns in the GDPR column set
  var rusGdprEndCol = 9;
  //creating a 1d arr of rus application IDs
  var rusOned = flatten_arr(rusArray);
  //removes the last item of the array since it is a blank item
  rusOned.pop();
  //finding the original information looping and pasting
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
  ui.alert('Updates have been pulled.');
}

function push_updates() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  //rus sheet vars and functions
  var rus = ss.getSheetByName('Review Updates');
  var rusLr = rus.getLastRow();
  var rusLc = rus.getLastColumn();
  var rusRange = rus.getRange(2, 1, rusLr, 1);
  var rusArray = rusRange.getValues();
  var rusStartRow = rusRange.getRow();
  var rusTitleColumnArr =rus.getRange(1, 1, 1, rusLc).getValues();
  //find gdpr section start column position
  var rustcaOned = flatten_arr(rusTitleColumnArr);
  var rusGdprBegCol = find_col(rustcaOned, 'GDPR Data (Y,N)');
  var rusGdprEndCol = 9;
  //1D array of rus application IDs
  var rusOned = flatten_arr(rusArray);

  //1D array of TPSL header column values and find gdpr section start column
  var tcaOned = flatten_arr(tpslTitleColumnArr);
  var tpslGdprBegCol = find_col(tcaOned, 'GDPR Data (Y,N)');
  var tpslGdprEndCol = 9;

  //1D arr of tpsl application IDs
  var tpslOned = flatten_arr(tpslArray);

  //create an array of the updates column and find row position and delete the ones that dont equal y
  //the updates column needs to be the last column in the sheet
  var rusLr2 = rus.getLastRow();
  var rusLcIsUpdatesCol = rus.getLastColumn();
  var rusUpdatesArr = rus.getRange(2, rusLcIsUpdatesCol, rusLr2, 1).getValues();

  var rusUpdatesArrOned = flatten_arr(rusUpdatesArr);


  // BEGINNING OF LOOK UP AND REPLACE
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
