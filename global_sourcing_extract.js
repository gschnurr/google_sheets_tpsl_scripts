function gs_extract() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var gsExtColFilArr = ['SL-ID', 'Application', 'Supplier (Third Party Vendor)', 'Application Manager', 'Functional Description',
  'Last Notice Period', 'Agreement End Date', 'Agreement Transition Status', 'Yearly Agreement Value (SEK)', 'Comments'];
  //create extract sheet and name
  spreadsheet.insertSheet(2);
  spreadsheet.getActiveSheet().setName('Global Services New Extract');

  var gsExtract = ss.getSheetByName('Global Services New Extract');
  //paste value of specified range
  var tpslAllCellsData = tpslAllCells.getValues();
  var tpslAllNotes = tpslAllCells.getNotes();
  var tpslAllDvRules = tpslAllCells.getDataValidations();

  gsExtract.getRange(1, 1, tpslLr, tpslLc).setValues(tpslAllCellsData);
  var gsExtractLc = gsExtract.getLastColumn();
  var gsExtractLr = gsExtract.getLastRow();
  gsExtract.getRange(1, 1, gsExtractLr, gsExtractLc).setNotes(tpslAllNotes);
  gsExtract.getRange(1, 1, gsExtractLr, gsExtractLc).setDataValidations(tpslAllDvRules);
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

  var gsExtractIdColPos = find_col(gsExtractOned, 'SL-ID');
  var gsExtractLnpColPos = find_col(gsExtractOned, 'Last Notice Period');
  var gsExtractAedColPos = find_col(gsExtractOned, 'Agreement End Date');
  var gsExtractAtsColPos = find_col(gsExtractOned, 'Agreement Transition Status');
  var gsExtractYavColPos = find_col(gsExtractOned, 'Yearly Agreement Value');

  var gsExtractIdArray = gsExtract.getRange(2, gsExtractIdColPos, gsExtractLr, 1).getValues();
  var gsExtractIdOned = flatten_arr(gsExtractIdArray);
  var gsExtractContractDataExistsArr = []; //boolean array of true or false answering if values exist?

  for (var vv = 0; vv < gsExtractIdOned.length; vv++) {
    var appRow = vv + 1;
    var lastNotPerVal = gsExtract.getRange(appRow, gsExtractLnpColPos, 1, 1).getValue();
    var agrEndDatVal = gsExtract.getRange(appRow, gsExtractAedColPos, 1, 1).getValue();
    var agrTranStaVal = gsExtract.getRange(appRow, gsExtractAtsColPos, 1, 1).getValue();
    var yeaAgrVal = gsExtract.getRange(appRow, gsExtractYavColPos, 1, 1).getValue();

    if (agrTranStaVal != '') {
      gsExtractContractDataExistsArr.push('true');
    }
    else if (agrEndDatVal != '') {
      gsExtractContractDataExistsArr.push('true');
    }
    else if (yeaAgrVal != '') {
      gsExtractContractDataExistsArr.push('true');
    }
    else if (lastNotPerVal != '') {
      gsExtractContractDataExistsArr.push('true');
    }
    else {
      gsExtractContractDataExistsArr.push('false');
    }
  }
  var gsExtRowFilArr = ['true'];

  filter_rows(gsExtractContractDataExistsArr, gsExtRowFilArr, gsExtract);

  gsExtract.copyTo('1bz7mzCqrBJNfNsAl7FPsUJ6dgov703mTanfqIoywC-4');

};



function gs_extract_data_verification() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var gseNew = ss.getSheetByName('Global Services New Extract');
  var gseNewLr = gseNew.getLastRow();
  var gseNewLc = gseNew.getLastColumn();
  var gseNewTca = gseNew.getRange(1, 1, 1, gseNewLc).getValues();
  var gseNewTconed = flatten_arr(gseNewTca);
  var gseNewIdColPos = find_col(gseNewTconed, 'SL-ID');
  var gseNewIdArr = gseNew.getRange(2, gseNewIdColPos, gseNewLr, 1).getValues();
  var gseNewIdOned = flatten_arr(gseNewIdArr);

  var gseOld = ss.getSheetByName('Global Services Last Extract');
  var gseOldLr = gseOld.getLastRow();
  var gseOldLc = gseOldLc.getLastColumn();
  var gseOldTca = gseOld.getRange(1, 1, 1, gseOldLc).getValues();
  var gseOldTconed = flatten_arr(gseOldTca);
  var gseOldIdColPos = find_col(gseOldTconed, 'SL-ID');
  var gseOldIdArr = gseOld.getRange(2, gseOldIdColPos, gseOldLr, 1).getValues();
  var gseOldIdOned = flatten_arr(gseOldIdArr);

//beginning of validating the new data against the old data by looping through the new data IDs
  for (var ln = 0; ln < gseNewIdOned.length; ln++) {
    var gseNewId = gseNewIdOned[ln];
    var gseNewIdRow = ln + 2;

    if (gseOldIdOned.indexOf(gseNewId) > -1) {
      var gseOldIdIndex = gseOldIdOned.indexOf(gseNewId);
      var gseOldIdRow = gseOldIdIndex + 2;

      for (var md = 0; md < gseNewTconed.length; md++) {
        var colTtm = gseNewTconed[md];
        var gseNewColPos = md + 1;

        if (gseOldTconed.indexOf(colTtm) > -1) {
          var gseColTitleIndex = gseOldTconed.indexOf(colTtm);
          var gseOldColPos = gseColTitleIndex + 1;
          var gseOldVal = gseOld.getRange(gseOldIdRow, gseOldColPos, 1, 1).getValue();
          var gseNewCellPos = gseNew.getRange(gseNewIdRow, gseNewColPos, 1, 1);
          var gseNewVal = gseNewCellPos.getValue();

          if (gseNewVal == gseOldVal) {
            continue;
          }
          else {
            gseNewCellPos.setBackground('yellow');
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
// need to change the name of the new sheet that is to be delivered then create a copy in its own workbook
//create an email and add that new copy as attachment to the email with the email body explaining colors
// need to change the name of the old extract we compared against to be archived something then change the latest one to the old name

}//function end



















//bullshit comment delete lolz=
