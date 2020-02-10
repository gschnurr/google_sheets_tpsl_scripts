function pp_push_updates_wrapper() {
  var ui = SpreadsheetApp.getUi();

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
  var ui = SpreadsheetApp.getUi();
  var spreadsheetName = spreadsheet.getName();
  var sheetName = ss.getSheetName();

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
