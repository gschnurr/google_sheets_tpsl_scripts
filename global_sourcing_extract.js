//3 Copy this into a specific workbook for this purpose as a new sheet. From there run the check against the older version.

//4 check function should be in all extracts but will copy over the old vs new in the same workbook
//  use the old tpsl macro way to check for changes + we will also need to check if there are any new SL-IDS entirely.

function gs_extract() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var gsExtColFilArr = ['SL-ID', 'Application', 'Supplier (Third Party Vendor)', 'Application Manager', 'Functional Description',
  'Last Notice Period', 'Agreement End Date', 'Agreement Transition Status', 'Yearly Agreement Value (SEK)', 'Comments'];
  //create extract sheet and name
  spreadsheet.insertSheet(2);
  spreadsheet.getActiveSheet().setName('Global Services Extract');

  var gsExtract = ss.getSheetByName('Global Services Extract');
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

};
