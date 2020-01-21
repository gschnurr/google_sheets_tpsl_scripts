function onEdit(e){
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var tcaOned = flatten_arr(tpslTitleColumnArr);
  var modDateColPos = find_col(tcaOned, 'Last Modified Date');
  var editRange = e.range;
  var editRowPos = editRange.getRow();
  var editColPos = editRange.getColumn();
  var editColTitle = tpsl.getRange(2, editColPos, 1, 1).getValue();
  if (editRowPos > 3 && editColTitle != 'Last Modified Date') {
    tpsl.getRange(editRowPos, modDateColPos, 1, 1).setValue(date);
  }
}
