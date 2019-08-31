
//Sheet
var ss = SpreadsheetApp.getActiveSpreadsheet();
var tpsl = ss.getSheetByName('1_Business Systems');

//Data
var tpslLc = tpsl.getLastColumn();

//User
var authorizedUsers = ['gibson.schnurr@izettle.com', 'linn.andersson@izettle.com', 'josefin.eklund@izettle.com', 'maaike.gerritse@izettle.com', 'markus.kanerva@izettle.com', 'roxanne.baumann@izettle.com', 'shumel.rahman@izettle.com'];
var currentUser = Session.getActiveUser().getEmail();
var numberOfAuthUsers = (authorizedUsers.length - 1);

//UI
var ui = SpreadsheetApp.getUi();
