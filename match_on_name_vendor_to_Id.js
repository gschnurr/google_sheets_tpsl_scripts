function match_app_info() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var tcaOned = flatten_arr(tpslTitleColumnArr);
  var tpslAppNameColPos = find_col(tcaOned, 'Application');
  var tpslAppNameArr = tpsl.getRange(2, tpslAppNameColPos, tpslLr, 1).getValues();
  var tpslAppNameOned = flatten_arr(tpslAppNameArr);
  var tpslSupNameColPos = find_col(tcaOned, 'Supplier (Third Party Vendor)');
  var tpslSupNameArr = tpsl.getRange(2, tpslSupNameColPos, tpslLr, 1).getValues();

//call ppe cols array

  var importSheet = ss.getSheetByName('TPSL Import');
  var ppe = ss.getSheetByName('PayPal Extract');

  var iSLr = importSheet.getLastRow();
  var iSLc = importSheet.getLastColumn();
  var iSTcArr = importSheet.getRange(1, 1, 1, iSLc).getValues();
  var iSTcOned = flatten_arr(iSTcArr);
  // what we match on is Below
  var iSAppNameColPos = find_col(iSTcOned, 'Application');
  var iSAppNameArr = importSheet.getRange(2, iSAppNameColPos, iSLr, 1).getValues();
  var iSAppNameOned = flatten_arr(iSAppNameArr);
  var iSSupNameColPos = find_col(iSTcOned, 'Supplier (Third Party Vendor)');
  var iSSupNameArr = importSheet.getRange(2, iSSupNameColPos, iSLr, 1).getValues();

  var noAppNameMatchArr = [];

  for (var dso = 0; dso < iSAppNameOned.length; dso++) {
    if (tpslAppNameOned.indexOf(iSAppNameOned[dso]) > -1) {
      var tpslAppNameRow = (tpslAppNameOned.indexOf(iSAppNameOned[dso]) + 2);
      var iSAppNameRow = dso + 2;
      var appName = iSAppNameOned[dso];
      logs_tst('There was a match based on Application Name, ' + iSAppNameOned[dso]);
      for (var tcm = 0; tcm < ppeColsArr.length; tcm++) {
        var iSLc = importSheet.getLastColumn();
        var resetIsTcArr = importSheet.getRange(1, 1, 1, iSLc).getValues();
        var resetIsTcOned = flatten_arr(resetIsTcArr);
        if (resetIsTcOned.indexOf(ppeColsArr[tcm]) > -1) {
          var iSCurColPos = find_col(resetIsTcOned, ppeColsArr[tcm]);
          var tpslCurColPos = find_col(tcaOned, ppeColsArr[tcm]);
          var iSCurCell = importSheet.getRange(iSAppNameRow, iSCurColPos, 1, 1);
          var iSCurValue = iSCurCell.getValue();
          if (iSCurValue != '') {
            logs_tst('Data already exists in the import sheet for the item of Applicaiton ' + appName + '. In column ' + ppeColsArr[tcm]);
            continue;
          }
          else {
            // insert data in this column from the tpsl
            var tpslValueForCurCell = tpsl.getRange(tpslAppNameRow, tpslCurColPos, 1, 1).getValue();
            iSCurCell.setValue(tpslValueForCurCell);
            continue;
          }
        }
        else {
          // what to do if the column does not exist in the import sheet
          var arrColNum = tcm + 1;
          var colInsertBeforePos = arrColNum;
          var tpslCurColPos = find_col(tcaOned, ppeColsArr[tcm]);
          importSheet.insertColumnBefore(colInsertBeforePos);
          importSheet.getRange(1, arrColNum, 1, 1).setValue(ppeColsArr[tcm]);
          var newCellPos = importSheet.getRange(iSAppNameRow, arrColNum, 1, 1);
          var tpslValueForNewCell = tpsl.getRange(tpslAppNameRow, tpslCurColPos, 1, 1).getValue();
          newCellPos.setValue(tpslValueForNewCell);
          continue;
        }
      } // closing of the title column seach loop
    }
    else {
      //what to do if we do not match on app name
      noAppNameMatchArr.push(iSAppNameOned[dso]);
    }
  } // closing of the app name loop
  var iSLc = importSheet.getLastColumn();
  importSheet.insertColumnAfter(iSLc);
  importSheet.getRange(1, (iSLc + 1), 1, 1).setValue('Updates? (Y/N) If yes please make the updates in this sheet');
  var compLogs = Logger.getLog();
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'Import into TPSL Logs', compLogs);
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'Apps without a Match on Import', noAppNameMatchArr);
}//function end
