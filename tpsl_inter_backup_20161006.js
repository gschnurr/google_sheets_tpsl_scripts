/** @OnlyCurrentDoc */
//Sheet
var spreadsheet = SpreadsheetApp.getActive();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
var spreadsheetName = spreadsheet.getName();
var sheetName = ss.getSheetName();
var tpsl = ss.getSheetByName('1_Business Systems');
var expGen = ss.getSheetByName('Export Generator');

//Time
var tz = ss.getSpreadsheetTimeZone();
var date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
var exportName = 'Generic Export ' + date;

//Data
var tpslLc = tpsl.getLastColumn();
var tpslLr = tpsl.getLastRow();
var tpslAllCells = tpsl.getRange(1, 1, tpslLr, tpslLc);
var tpslTitleColumnArr = tpsl.getRange(2, 1, 1, tpslLc).getValues();
//the tpslrange and array creating an arr of the application IDs
var tpslRange = tpsl.getRange(4, 1, tpslLr, 1);
var tpslArray = tpslRange.getValues();
var tpslStartRow = tpslRange.getRow();
var expGenLr = expGen.getLastRow();
var expGenColumnArr = expGen.getRange(3, 1, expGenLr, 1).getValues();
//This array contains all of the columns that you want to keep in the extract
//If you would like a new column added please add the column header exactly as it is into the array
var ppeColsArr = ['SL-ID', 'Application', 'Supplier (Third Party Vendor)', 'Application Manager',
'Business System Owner', 'GDPR Data (Y,N)', 'Employee Data', 'End Customer Data', 'Merchant Data',
'Vendor Category', 'Purpose', 'Data Disclosed', 'Data shared with third party? (Y,N,N/A)',
'Headquarter location'];

var gdprWizTColArr = ['Business System Owner', 'GDPR Data (Y,N)', 'Employee Data', 'End Customer Data', 'Merchant Data',
'Vendor Category', 'Purpose', 'Data Disclosed', 'Data shared with third party? (Y,N,N/A)',
'Headquarter location'];

//User
var authorizedUsers = ['gibson.schnurr@izettle.com', 'linn.andersson@izettle.com',
'josefin.eklund@izettle.com', 'maaike.gerritse@izettle.com', 'markus.kanerva@izettle.com',
'roxanne.baumann@izettle.com', 'shumel.rahman@izettle.com'];
var currentUser = Session.getActiveUser().getEmail();
var numberOfAuthUsers = (authorizedUsers.length - 1);

//UI
var ui = SpreadsheetApp.getUi();

//Global Functions
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

function find_col(tleColFlatArr, colToFind) {
  var colPos;
  for (var i = 0; i < tleColFlatArr.length; i++) {
    if (tleColFlatArr[i] == colToFind) {
      var colPos = i + 1; //add one because arrays start at 0 not 1
    }
    else {
      continue;
    }
  }
  return colPos;
}

function onOpen(e) {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //Simple Exports Menu Item
    ui.createMenu('Simple Exports')
      .addItem('Generic Export Instructions', 'gen_export_ins')
      .addSeparator()
      .addItem('Generate Export', 'gen_export_wrapper')
      .addSeparator()
      .addItem('Clean Export', 'clean_export_wrapper')
      .addToUi();
  //PP Export Menu Item
    ui.createMenu('PP Exports')
      .addItem('PP Quarterly Macro Instructions', 'pp_information_prompt')
      .addSeparator()
      .addItem('PayPal GDPR Extract', 'tpsl_pp_extract')
      .addSeparator()
      .addItem('Get Updates', 'get_updates')
      .addSeparator()
      .addItem('Push Updates', 'pp_push_updates_wrapper')
      .addToUi();
    var menu = [{name: 'Set up conference', functionName: 'setUpConference_'}];
    SpreadsheetApp.getActive().addMenu('Conference', menu);
  //Checking if clean export should be run which is defined first by spreadsheet name then by number sheets
  if (spreadsheetName == exportName && sheets.length > 2) {
    clean_export_wrapper();
  }
};


function pp_gdpr_wizard_wrapper(){
  if (sheetName == 'PayPal Extract') {
    pp_gdpr_wizard();
  }
  else {
    ui.alert('You can not run the GDPR Wizard in this sheet, please move to the appopriate sheet to run the wizard.');
  }
}


function pp_push_updates_wrapper() {
  for (var i = 0; i < authorizedUsers.length; i++) {
    if (currentUser == authorizedUsers[i]) {
      var response = ui.alert('WARNING: Google Sheets is set to run the push_updates macro. This macro will overwrite existing data in the TPSL document. Are sure you want to continue with this Macro?', ui.ButtonSet.YES_NO);
      if (response == ui.Button.NO) {
        return;
      }
      push_updates();
      MailApp.sendEmail('gibson.schnurr@izettle.com',
                'TPSL PP Updates Push',
                'The push updates macro has been run. The running user was ' + currentUser + '.');
      ui.alert('Updates have been pushed.');
      break;
    }
    else if (i < numberOfAuthUsers) {
      continue;
    }
    else if (i == numberOfAuthUsers) {
      MailApp.sendEmail('gibson.schnurr@izettle.com',
                'Unauthorized Macro Attempt - Push Updates',
                currentUser + ' attempted to run the push updates macro.');
      ui.alert('ERROR: You are not listed as an authorized user of this macro. Please contact Gibson to add you to the list of authorized users');
    }
  }
}

function gen_export_wrapper() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  //Checking if any columns have been chosen for export
  //it does not matter if the same column has been chosen multiple times
  var expGenOned = flatten_arr(expGenColumnArr);
  var numExpCols = 0;
  for (var i = 0; i < expGenOned.length; i++) {
    if (expGenOned[i] != '') {
      ++numExpCols;
      continue;
    }
    else if (expGenOned[i] == '' && i != (expGenOned.length - 1)) {
      continue;
    }
    else if (i == (expGenOned.length - 1) && numExpCols != 0) {
      gen_export();
      MailApp.sendEmail('gibson.schnurr@izettle.com',
                'General Export Created',
                'The general export macro was run successfully. The running user was ' + currentUser + '.' +
                'The exported columns are ' + expGenOned);
      ui.alert('Export Created, Please check your google sheet files for the Generic Export with Todays Date.');
    }
    else if (i == (expGenOned.length - 1) && numExpCols == 0) {
      ui.alert('Whoops! You have not selected any columns to export. Please select at least one column.');
      expGen.activate();
    }
    else {
      MailApp.sendEmail('gibson.schnurr@izettle.com',
                'General Export Complete Failure',
                'The general export wrapper failed to run the macro and was forced to the final else, which should not happen under any reasonable explanation. The running user was ' + currentUser + '.');
      ui.alert('Whoops something went wrong! Your macro administrator has been notified via email. Please confirm that your export columns have values.');
    }
  }
}

function clean_export_wrapper() {
  if (spreadsheetName == 'TPSL 2.0') {
    SpreadsheetApp.getUi().alert('Error: this macro cannot be run in the master TPSL document, please run this only in your generated export.');
    return;
  }
  var response = ui.alert('WARNING: Google Sheets is set to run the clean_export macro. Confirm you are not in the Master TPSL document, this macro should only be run in generated exports. Are sure you want to continue with this Macro?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {
    return;
  }
  clean_export();
  MailApp.sendEmail('gibson.schnurr@izettle.com',
            'General Export Cleaned',
            'The clean export macro was run on ' + spreadsheetName + '. The running user was ' + currentUser + '.');
}

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
  ui.showSidebar(htmlOutput);
}

function tpsl_pp_extract() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var tcaOned = flatten_arr(tpslTitleColumnArr);
  //finding gpdr column start position
  var gdprColPos = find_col(tcaOned, 'GDPR Data (Y,N)');
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

  //paste value of specified range this should probably change from a paste function to a set function
  ppe.getRange('A1').activate();
  tpslAllCells.copyTo(ppe.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  var tpslAllNotes = tpslAllCells.getNotes();
  var tpslAllDvRules = tpslAllCells.getDataValidations();
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  var ppeAllCells = ppe.getRange(1, 1, ppeLr, ppeLc);
  ppeAllCells.setNotes(tpslAllNotes);
  ppeAllCells.setDataValidations(tpslAllDvRules);


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
  ui.alert('Extract Created, Please check your google sheet files for the PayPal Extract with Todays Date');
};

function get_updates() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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

function gen_export_ins() {
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
  for (var i=0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() != 'Generic Export') {
      sheets[i].activate();
      spreadsheet.deleteActiveSheet();
    }
  }
}

function pp_wizard_information_prompt() {
  var htmlOutput = HtmlService
    .createHtmlOutput ('<p>For this review you should review the GDPR information for the applications that you are the assigned business system owner. You can update the information or use the Update Wizard. To use the Wizard see the steps below, the purpose of this information.</p>' +
    '<p> Purpose: Type up the purpose of this update exercise.</p>' +
    '<p> 1) To use the Update Wizard start the wizard by going to the GDPR Update Wizard Menu item and select Run Wizard.</p>' +
    '<p> 2) The wizard will cycle through each application that you are the assigned business owner. You will be presented with a prompt for each column.</p>' +
    '<p> 3) This prompt will contain the Application Name, Column Title, Column Description, and Current Information.</p>' +
    '<p> 4) If the current information is correct press the cancel button.</p>' +
    '<p> 5) If the current information is incorrect type the correct information in the free text field and press OK. The cell will automatically be updated with the new information.</p>' +
    '<p> 7) If you would like to skip to the next application type "next application" in the free text field and press OK.</p>' +
    '<p> 8) If you would like to exit the the wizard completely type "ESCAPE" in the free text field and press OK. This is case sensative.</p>'
  )
    .setTitle('GDPR Update Wizard Instructions');
  ui.showSidebar(htmlOutput);
}

function pp_gdpr_wizard() {
  pp_wizard_information_prompt();

  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ppe = ss.getSheetByName('PayPal Extract');
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  var ppeTitleColumnArr = ppe.getRange(1, 1, 1, ppeLc).getValues();
  var currentWizUser = Session.getActiveUser().getEmail();

  var ppeOned = flatten_arr(ppeTitleColumnArr);

  var busSysOwnColPos = find_col(ppeOned, 'Business System Owner');
  var busSysOwnerArr = ppe.getRange(2, busSysOwnColPos, ppeLr, 1).getValues();
  var bsoOned = flatten_arr(busSysOwnerArr);

  var appNameColPos = find_col(ppeOned, 'Application');
  var appNameArr = ppe.getRange(2, appNameColPos, ppeLr, 1).getValues();
  var anOned = flatten_arr(appNameArr);

  //this is an array of the row number of the applications owned by the currentWizUser
  var wizUserAppsRowNumArr = [];
  var wizUserAppsArr = [];
  for (var f = 0; f < bsoOned.length; f++) {
    if (bsoOned[f] == currentWizUser) {
      var cwuAppRow = f + 2;
      wizUserAppsRowNumArr.push(cwuAppRow);
      var appNameCell = ppe.getRange(cwuAppRow, appNameColPos, 1, 1).getValue();
      wizUserAppsArr.push(appNameCell);
    }
    else {
      continue;
    }
  }
  //PPE cols doesnt work
  //creates an arr of all col pos of columns to be updated
  var gdprWizColPosArr = [];
  for (var c = 0; c < gdprWizTColArr.length; c++) {
    var gdprWizColPos = find_col(ppeOned, gdprWizTColArr[c]);
    gdprWizColPosArr.push(gdprWizColPos);
  }

  for (var u = 0; u < wizUserAppsRowNumArr.length; u++) {
    for (var v = 0; v < gdprWizColPosArr.length; v++) {
      var appToUpdate = wizUserAppsArr[u];
      var colTitle = gdprWizTColArr[v];
      var colDescription = ppe.getRange(1, gdprWizColPosArr[v], 1, 1).getNotes();
      ppe.getRange(wizUserAppsRowNumArr[u], appNameColPos, 1, 1).activateAsCurrentCell();
      var currentInfo = ppe.getRange(wizUserAppsRowNumArr[u], gdprWizColPosArr[v], 1, 1).getValue();
      var currentInfoCheck;
      if (currentInfo == '') {
        currentInfoCheck = '<BLANK>';
      }
      else {
        currentInfoCheck = currentInfo;
      }


      var wizUserResp = Browser.inputBox(appToUpdate,
        'Please update the information relating to the ' + colTitle + ' field' +
        '. The description of this field is: ' + colDescription +
        ' The current information is '+ currentInfoCheck +
        '. Type your updates in the response field and press OK to accept the changes. ' +
        'If you do not have any changes to make press cancel to move to the next field for review. ',
        Browser.Buttons.OK_CANCEL);
      if (wizUserResp == 'cancel') {
        continue;
      }
      else if (wizUserResp == '') {
        continue;
      }
      else if (wizUserResp == 'next application') {
        break;
      }
      else if (wizUserResp == 'ESCAPE') {
        return;
      }
      else {
        ppe.getRange(wizUserAppsRowNumArr[u], gdprWizColPosArr[v], 1, 1).setValue(wizUserResp);
      }
    }
  }
}



function test_gen_form() {

  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ppe = ss.getSheetByName('PayPal Extract');
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  var ppeTitleColumnArr = ppe.getRange(1, 1, 1, ppeLc).getValues();
  var currentWizUser = Session.getActiveUser().getEmail();

  var ppeOned = flatten_arr(ppeTitleColumnArr);

  var busSysOwnColPos = find_col(ppeOned, 'Business System Owner');
  var busSysOwnerArr = ppe.getRange(2, busSysOwnColPos, ppeLr, 1).getValues();
  var bsoOned = flatten_arr(busSysOwnerArr);

  var appNameColPos = find_col(ppeOned, 'Application');
  var appNameArr = ppe.getRange(2, appNameColPos, ppeLr, 1).getValues();
  var anOned = flatten_arr(appNameArr);

  //this is an array of the row number of the applications owned by the currentWizUser
  var wizUserAppsRowNumArr = [];
  var wizUserAppsArr = [];
  for (var f = 0; f < bsoOned.length; f++) {
    if (bsoOned[f] == currentWizUser) {
      var cwuAppRow = f + 2;
      wizUserAppsRowNumArr.push(cwuAppRow);
      var appNameCell = ppe.getRange(cwuAppRow, appNameColPos, 1, 1).getValue();
      wizUserAppsArr.push(appNameCell);
    }
    else {
      continue;
    }
  }
  //creates an arr of all col pos of columns to be updated
  var gdprWizColPosArr = [];
  for (var c = 0; c < gdprWizTColArr.length; c++) {
    var gdprWizColPos = find_col(ppeOned, gdprWizTColArr[c]);
    gdprWizColPosArr.push(gdprWizColPos);
  }

  var userUpdatesForm = FormApp.create(currentWizUser + 'applications');
    for (var u = 0; u < wizUserAppsRowNumArr.length; u++) {
    for (var v = 0; v < gdprWizColPosArr.length; v++) {
      var appToUpdate = wizUserAppsArr[u];
      var colTitle = gdprWizTColArr[v];
      var colDescription = ppe.getRange(1, gdprWizColPosArr[v], 1, 1).getNotes();
      ppe.getRange(wizUserAppsRowNumArr[u], appNameColPos, 1, 1).activateAsCurrentCell();
      var currentInfo = ppe.getRange(wizUserAppsRowNumArr[u], gdprWizColPosArr[v], 1, 1).getValue();
      var currentInfoCheck;
      if (currentInfo == '') {
        currentInfoCheck = '<BLANK>';
      }
      else {
        currentInfoCheck = currentInfo;
      }

      userUpdatesForm.addTextItem()
        .setTitle(appToUpdate + "'s " +colTitle)
        .setHelpText('. The description of this field is: ' + colDescription + ' The current information is '+ currentInfoCheck + '. Type your updates in the response field and press OK to accept the changes. ' + 'If you do not have any changes to make press cancel to move to the next field for review. ');
      }
    }
}

var loggingOnOff = 'OFF';

function logger_wrapper_test(logContent) {
  if (loggingOnOff == 'OFF') {
  }
  else if (loggingOnOff == 'ON') {
    Logger.log(logContent);
  }
  else {
  }
}

function testlogwrapper() {
 var testvar = 'Gibson';

 logger_wrapper_test('This is a test log did it work? ' + '-' + testvar);
}
