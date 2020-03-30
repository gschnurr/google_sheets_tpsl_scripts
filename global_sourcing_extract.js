function gs_extract() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var gsExtColFilArr = ['SL-ID', 'Application', 'Supplier (Third Party Vendor)', 'Application Manager',
  'Last Notice Period', 'Agreement End Date', 'Agreement Transition Status', 'Approved Budget Value (SEK)'];
  //create extract sheet and name
  spreadsheet.insertSheet(2);
  spreadsheet.getActiveSheet().setName('Global Sourcing New Extract');

  var gsExtract = ss.getSheetByName('Global Sourcing New Extract');
  //paste value of specified range
  var tpslAllCellsData = tpslAllCells.getValues();
  var tpslAllNotes = tpslAllCells.getNotes();

  gsExtract.getRange(1, 1, tpslLr, tpslLc).setValues(tpslAllCellsData);
  var gsExtractLc = gsExtract.getLastColumn();
  var gsExtractLr = gsExtract.getLastRow();
  gsExtract.getRange(1, 1, gsExtractLr, gsExtractLc).setNotes(tpslAllNotes);
  gsExtract.deleteRow(3);
  gsExtract.deleteRow(1);
  //Copying format of title row from tpsl to ppe
  gsExtract.getRange('A1').activate();
  tpsl.getRange(2, 1, 1, tpslLc).copyTo(gsExtract.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  //freezing title row of ppe
  gsExtract.setFrozenRows(1);
  var gsExtractLc = gsExtract.getLastColumn();
  //deleting un-needed columns
  var gsExtractColumnArr = gsExtract.getRange(1, 1, 1, gsExtractLc).getValues();
  var gsExtractOned = flatten_arr(gsExtractColumnArr);

  filter_cols(gsExtractOned, gsExtColFilArr, gsExtract);

  var gsExtractLc = gsExtract.getLastColumn();
  var gsExtractLr = gsExtract.getLastRow();

  var gsExtractColumnArr = gsExtract.getRange(1, 1, 1, gsExtractLc).getValues();
  var gsExtractOned = flatten_arr(gsExtractColumnArr);

  var gsExtractIdColPos = find_col(gsExtractOned, 'SL-ID');
  var gsExtractLnpColPos = find_col(gsExtractOned, 'Last Notice Period');
  var gsExtractAedColPos = find_col(gsExtractOned, 'Agreement End Date');
  var gsExtractAtsColPos = find_col(gsExtractOned, 'Agreement Transition Status');
  var gsExtractYavColPos = find_col(gsExtractOned, 'Approved Budget Value (SEK)');

  var gsExtractIdArray = gsExtract.getRange(2, gsExtractIdColPos, gsExtractLr, 1).getValues();
  var gsExtractIdOned = flatten_arr(gsExtractIdArray);
  var gsExtractContractDataExistsArr = []; //boolean array of true or false answering if values exist?
  var testingArr = []

  for (var vv = 0; vv < gsExtractIdOned.length; vv++) {
    var appRow = vv + 2;
    var lastNotPerVal = gsExtract.getRange(appRow, gsExtractLnpColPos, 1, 1).getValue();
    var agrEndDatVal = gsExtract.getRange(appRow, gsExtractAedColPos, 1, 1).getValue();
    var agrTranStaVal = gsExtract.getRange(appRow, gsExtractAtsColPos, 1, 1).getValue();
    logs_tst('value to check ' + agrTranStaVal);
    var yeaAgrVal = gsExtract.getRange(appRow, gsExtractYavColPos, 1, 1).getValue();

    if (agrTranStaVal != '') {
      gsExtractContractDataExistsArr.push('true');
      testingArr.push(['true', agrTranStaVal, gsExtractIdOned[vv]]);
      continue;
    }
    else {
      gsExtractContractDataExistsArr.push('false');
      testingArr.push(['false', agrTranStaVal]);
      continue;
    }
  }
  var gsExtRowFilArr = ['true'];
  logs_tst(testingArr);

  filter_rows(gsExtractContractDataExistsArr, gsExtRowFilArr, gsExtract);
  var copyDestination = SpreadsheetApp.openById('1bz7mzCqrBJNfNsAl7FPsUJ6dgov703mTanfqIoywC-4')
  gsExtract.copyTo(copyDestination);
  spreadsheet.deleteSheet(gsExtract);
};



function gs_extract_data_verification() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var gseChangeName = ss.getSheetByName('Copy of Global Sourcing New Extract');
  gseChangeName.setName('Global Sourcing New Extract');
  var gseNew = ss.getSheetByName('Global Sourcing New Extract');
  var gseNewLr = gseNew.getLastRow();
  var gseNewLc = gseNew.getLastColumn();
  var gseNewTca = gseNew.getRange(1, 1, 1, gseNewLc).getValues();
  var gseNewTconed = flatten_arr(gseNewTca);
  var gseNewIdColPos = find_col(gseNewTconed, 'SL-ID');
  var gseNewIdArr = gseNew.getRange(2, gseNewIdColPos, gseNewLr, 1).getValues();
  var gseNewIdOned = flatten_arr(gseNewIdArr);

  var gseOld = ss.getSheetByName('Global Sourcing Last Extract');
  var gseOldLr = gseOld.getLastRow();
  var gseOldLc = gseOld.getLastColumn();
  var gseOldTca = gseOld.getRange(1, 1, 1, gseOldLc).getValues();
  var gseOldTconed = flatten_arr(gseOldTca);
  var gseOldIdColPos = find_col(gseOldTconed, 'SL-ID');
  var gseOldIdArr = gseOld.getRange(2, gseOldIdColPos, gseOldLr, 1).getValues();
  var gseOldIdOned = flatten_arr(gseOldIdArr);

//beginning of validating the new data against the old data by looping through the new data IDs
  for (var ln = 0; ln < gseNewIdOned.length; ln++) {
    var gseNewId = gseNewIdOned[ln];
    logs_tst('gseNewID = ' + gseNewId);
    var gseNewIdRow = ln + 2;
    logs_tst('gseNewIdRow = ' + gseNewIdRow);

    if (gseOldIdOned.indexOf(gseNewId) > -1) {
      var gseOldIdIndex = gseOldIdOned.indexOf(gseNewId);
      var gseOldIdRow = gseOldIdIndex + 2;
      logs_tst('gseOldIdRow = ' + gseOldIdRow);

      for (var md = 0; md < gseNewTconed.length; md++) {
        var colTtm = gseNewTconed[md];
        var gseNewColPos = md + 1;

        if (gseOldTconed.indexOf(colTtm) > -1) {
          var gseColTitleIndex = gseOldTconed.indexOf(colTtm);
          var gseOldColPos = gseColTitleIndex + 1;
          var gseOldCellPos = gseOld.getRange(gseOldIdRow, gseOldColPos, 1, 1);
          var gseNewCellPos = gseNew.getRange(gseNewIdRow, gseNewColPos, 1, 1);

          if (colTtm == 'Agreement End Date') {
            logs_tst('Date col found');
            var gseNewValDate = gseNewCellPos.getValue();
            var gseOldValDate = gseOldCellPos.getValue();

            var gseNewVal = gseNewValDate.valueOf();
            var gseOldVal = gseOldValDate.valueOf();
            logs_tst(gseOldColPos);
            logs_tst(gseOldVal);
          }
          else {
            var gseNewVal = gseNewCellPos.getValue();
            var gseOldVal = gseOldCellPos.getValue();
          }

          logs_tst('gseNewVal = ' + gseNewVal + ' gseOldVal = ' + gseOldVal);

          if (gseNewVal == gseOldVal) {
            continue;
          }
          else {
            gseNewCellPos.setBackground('yellow');
            gseOldCellPos.setBackground('orange');
          } //comparing values if statement
        } //finding matching column title if statement
        else {
          continue;
        }
      } //using already found approw data to find exact cells to validate for loop (TCA Loop)
    } //if statement for seeing if there is a matching ID or not in the old sheet from the new sheet
    else {
      gseNew.getRange(gseNewIdRow, 1, 1, gseNewLc).setBackground('green');
    } //else statement if the new extract ID does not exist in the old data this means its a new system
  } // for loop of new sheet IDs

//this for loop is checking if a system has been deleted from the tpsl since the last extract by checking if all the
//old ids exist in the new sheet if they dont it will add the old data to the bottom of the sheet and highlight red
  for (var lo = 0; lo < gseOldIdOned.length; lo++) {
    var gseMissingOldRow = lo + 2;
    if (gseNewIdOned.indexOf(gseOldIdOned[lo]) > -1) {
      continue;
    }
    else {
      var gseNewAddLr = gseNew.getLastRow();
      var gseNewEmptyRow = gseNewAddLr + 1;

      for (var ao = 0; ao < gseOldTconed.length; ao++) {
        var colTtm = gseOldTconed[ao];
        var gseMissingOldColPos = ao + 1;

        if (gseNewTconed.indexOf(colTtm) > -1) {
          var gseNewEmptyRowColPos = gseNewTconed.indexOf(colTtm) + 1;
          var gseMissingOldVal = gseOld.getRange(gseMissingOldRow, gseMissingOldColPos, 1, 1).getValue();
          gseNew.getRange(gseNewEmptyRow, gseNewEmptyRowColPos, 1, 1).setValue(gseMissingOldVal);
          gseNew.getRange(gseNewEmptyRow, gseNewEmptyRowColPos, 1, 1).setBackground('red');
        }
        else {
          continue;
        } // if index match closing on col pos
      } //for loop for col pos and replace closed
    } //else if there is no matching id value in new sheet
  } // closing of for loop looking for deleted applications
  var copyDestination = SpreadsheetApp.openById('1bz7mzCqrBJNfNsAl7FPsUJ6dgov703mTanfqIoywC-4')
  gseNew.copyTo(copyDestination);

  if (sheets.indexOf('iZettle Application Extract') > -1) {
    var deleteThis = ss.getSheetByName('iZettle Application Extract');
    ss.deleteSheet(deleteThis);
  }
  gseNew.setName('iZettle Application Extract');
  var gseCopy = ss.getSheetByName('Copy of Global Sourcing New Extract');
  gseOld.setName('GS Extract Archived on ' + date);
  gseCopy.setName('Global Sourcing Last Extract');

  var emailAttachment = DriveApp.getFileById('1bz7mzCqrBJNfNsAl7FPsUJ6dgov703mTanfqIoywC-4');
  GmailApp.createDraft('', 'iZettle Monthly Application Extract', 'Attached is a spreadsheet containing contracting information and statuses for iZettle third-party applications. All newly added applications will be highlighted in green, deleted applications will be highlighted in red, and changes to applicaiton information will be highlighted in yellow (with the previous data highlighted in orange in the archived sheet). Please let me know if you have any questions.', {
    attachments: [emailAttachment],
    name: 'Automatic Emailer Script'
  });
//create an email and add that new copy as attachment to the email with the email body explaining colors
}//function end
