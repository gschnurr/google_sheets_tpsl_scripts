function onEdit(e){
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();
  var currentEditSheet = ss.getSheetName();

  var editRange = e.range;
  var editRowPos = editRange.getRow();
  var editColPos = editRange.getColumn();
  var editUser = e.user;

  var tcaOned = flatten_arr(tpslTitleColumnArr);
  var tpslModDateColPos = find_col(tcaOned, 'Last Modified Date');
  var tpslEditUserColPos = find_col(tcaOned, 'Last Edit User');
  var tpslEditColTitle = tpsl.getRange(2, editColPos, 1, 1).getValue();

  var sheetsNameArr = [];
  for (var fs = 0; fs < sheets.length; fs ++) {
    var curSheetToAddToArr = sheets[fs].getSheetName();
    sheetsNameArr.push(curSheetToAddToArr);
  }

  if (sheetsNameArr.indexOf('PayPal Extract') > -1) {
    var ppe = ss.getSheetByName('PayPal Extract');
    var ppeLr = ppe.getLastRow();
    var ppeLc = ppe.getLastColumn();
    var ppeTitleColumnArr = ppe.getRange(1, 1, 1, ppeLc).getValues();
    var ppeTcaOned = flatten_arr(ppeTitleColumnArr);
    var ppeModDateColPos = find_col(ppeTcaOned, 'Last Modified Date');
    var ppeEditUserColPos = find_col(ppeTcaOned, 'Last Edit User');
    var ppeEditColTitle = ppe.getRange(1, editColPos, 1, 1).getValue();
  }

  if (editRowPos > 3 && tpslEditColTitle != 'Last Modified Date' && tpslEditColTitle != 'Last Edit User' && currentEditSheet == '1_Business Systems') {
    if (tpslEditColTitle == 'Last Notice Period' || tpslEditColTitle == 'Agreement End Date') {
      update_event(editRowPos);
    }
    else {
      tpsl.getRange(editRowPos, tpslModDateColPos, 1, 1).setValue(date);
      tpsl.getRange(editRowPos, tpslEditUserColPos, 1, 1).setValue(editUser);
    }
  }
  else if (editRowPos > 1 && ppeEditColTitle != 'Last Modified Date' && ppeEditColTitle != 'Last Edit User' && ppeEditColTitle != 'Updates? (Y/N) If yes please make the updates in this sheet' && currentEditSheet == 'PayPal Extract') {
    ppe.getRange(editRowPos, ppeModDateColPos, 1, 1).setValue(date);
    ppe.getRange(editRowPos, ppeEditUserColPos, 1, 1).setValue(editUser);
  }
}
