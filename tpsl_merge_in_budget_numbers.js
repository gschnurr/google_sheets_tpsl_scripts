function merge_finance_information() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var budg = ss.getSheetByName('APPROVED 2020 Budget');
  var budgLc = budg.getLastColumn();
  var budgLr = budg.getLastRow();
  //tpsl sheet is defined in global_variables
  var budgTca = budg.getRange(1, 1, 1, budgLc).getValues();
  var budgTcaOned = flatten_arr(budgTca);

  var budgIdColPos = find_col(budgTcaOned, 'SL-ID (new)');
  var budgTotSekColPos = find_col(budgTcaOned, 'Total SEK (including VAT)');

  var budgIdArr = budg.getRange(2, budgIdColPos, budgLr, 1).getValues();
  var budgIdOned = flatten_arr(budgIdArr);

  var tpslLc = tpsl.getLastColumn();
  var tpslLr = tpsl.getLastRow();

  var tcaOned = flatten_arr(tpslTitleColumnArr);

  var tpslIdColPos = find_col(tcaOned, 'SL-ID');
  var tpslYavColPos = find_col(tcaOned, 'Yearly Agreement Value (SEK)');

  var tpslIdArr = tpsl.getRange(4, tpslIdColPos, tpslLr, 1).getValues();
  var tpslIdOned = flatten_arr(tpslIdArr);

  for (var mf = 0; mf < tpslIdOned.length; mf++) {

    var tpslAppRow = mf + 4;
    var tpslId = tpslIdOned[mf];
    var tpslFinCell = tpsl.getRange(tpslAppRow, tpslYavColPos, 1, 1);

    if (budgIdOned.indexOf(tpslId) > -1 ) {
      var budgAppIndex = budgIdOned.indexOf(tpslId);
      var budgAppRow = budgAppIndex + 2;
      var budgMatId = budg.getRange(budgAppRow, budgIdColPos, 1, 1).getValue();
      logs_tst('Match found for TPSL ID = ' + tpslId + ' with BUDG ID = ' + budgMatId);
      var budgFinVal = budg.getRange(budgAppRow, budgTotSekColPos, 1, 1).getValue();
      tpslFinCell.setValue(budgFinVal);
      logs_tst('TPSL value replaced');
    }
    else {
      logs_tst('No match was found for ID = ' + tpslId);
      continue;
    }
  }
}
