/*
Framework:
wrapper
  check if the sheet is the PayPal Extract sheet if so check against running user


wizard
  get the user to find which applications they own
  make the application row and column one active cell for first application
  create pop-up dialog
    in the dialog contain a description of what the cell should contain
    show what is currently in the cell
    if it is a data validation cell contain the options and button set for those
    options
    otherwise have prompt for freetext which will replace the information
  then it will move to the next application

Another macro to be run by our team will run through the information and pull the
the updates based on a comparison to older data no longer based on the why


**/

// I need to switch everything taht says TPSL to PP Extract

function pp_gdpr_wizard() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var currentWizUser = Session.getActiveUser().getEmail();

  var tcaOned = flatten_arr(tpslTitleColumnArr);

  var busSysOwnColPos = find_col(tcaOned, 'Application Manager');
  var busSysOwnerArr = tpsl.getRange(4, busSysOwnColPos, tpslLr, 1).getValues();
  var bsoOned = flatten_arr(busSysOwnerArr);

  var appNameColPos = find_col(tcaOned, 'Application');
  var appNameArr = tpsl.getRange(4, appNameColPos, tpslLr, 1).getValues();
  var anOned = flatten_arr(appNameArr);

  //this is an array of the row number of the applications owned by the currentWizUser
  var wizUserAppsRowNumArr = [];
  var wizUserAppsArr = [];
  for (var f = 0; f < bsoOned.length; f++) {
    if (bsoOned[f] == currentWizUser) {
      var cwuAppRow = f + 4;
      wizUserAppsRowNumArr.push(cwuAppRow);
      var appNameCell = tpsl.getRange(cwuAppRow, appNameColPos, 1, 1).getValue();
      wizUserAppsArr.push(appNameCell);
    }
    else {
      continue;
    }
  }
  ui.alert(wizUserAppsArr);
  //PPE cols doesnt work
  //creates an arr of all col pos of columns to be updated
  var gdprWizColPosArr = [];
  for (var c = 0; c < gdprWizTColArr.length; c++) {
    var gdprWizColPos = find_col(tcaOned, gdprWizTColArr[c]);
    gdprWizColPosArr.push(gdprWizColPos);
  }

  for (var u = 0; u < wizUserAppsRowNumArr.length; u++) {
    for (var v = 0; v < gdprWizColPosArr.length; v++) {
      var appToUpdate = wizUserAppsArr[u];
      var colTitle = gdprWizTColArr[v];
      var colDescription = tpsl.getRange(2, gdprWizColPosArr[v], 1, 1).getNotes();
      tpsl.getRange(wizUserAppsRowNumArr[u], appNameColPos, 1, 1).activateAsCurrentCell();
      var currentInfo = tpsl.getRange(wizUserAppsRowNumArr[u], gdprWizColPosArr[v], 1, 1).getValue();
      var wizUserResp = Browser.inputBox(appToUpdate,
        'Please update the information relating to ' + colTitle + '. This means ' + colDescription +
        '. The current information is '+ currentInfo +
        '. When you are done press OK to accept changes or CANCEL to not make changes.',
        Browser.Buttons.OK_CANCEL);
      if (wizUserResp == 'cancel') {
        continue;
      }
      else {
        //code to replace current information
        
      }
    }
  }
}
