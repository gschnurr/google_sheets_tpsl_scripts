/** @OnlyCurrentDoc */

var tstloggingOnOff = 'ON';
var prdloggingOnOff = 'OFF';

//Sheet
var spreadsheet = SpreadsheetApp.getActive();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
var tpsl = ss.getSheetByName('1_Business Systems');
var expGen = ss.getSheetByName('Export Generator');

//Time
var tz = ss.getSpreadsheetTimeZone();
var date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
var exportName = 'Generic Export ' + date;

//Data
var tpslLc = tpsl.getLastColumn();
var tpslLr = tpsl.getLastRow();
var tpslAllCells = tpsl.getRange(1, 1, tpslLr, tpslLc);
var tpslTitleColumnArr = tpsl.getRange(2, 1, 1, tpslLc).getValues();
//the tpslrange and array creating an arr of the application IDs
var tpslRange = tpsl.getRange(4, 1, tpslLr, 1);
var tpslArray = tpslRange.getValues();
var tpslStartRow = tpslRange.getRow();
var expGenLr = expGen.getLastRow();
var expGenColumnArr = expGen.getRange(3, 1, expGenLr, 1).getValues();
//This array contains all of the columns that you want to keep in the extract
//If you would like a new column added please add the column header exactly as it is into the array
var ppeColsArr = ['SL-ID', 'Application', 'Supplier (Third Party Vendor)', 'Application Manager',
'Business System Owner', 'GDPR Data (Y,N)', 'Employee Data', 'End Customer Data', 'Merchant Data',
'Vendor Category', 'Purpose', 'Data Disclosed', 'Data shared with third party? (Y,N,N/A)',
'Headquarter location'];

var gdprBoTColArr = ['GDPR Data (Y,N)', 'Employee Data', 'End Customer Data', 'Merchant Data',
'Vendor Category', 'Purpose', 'Data Disclosed', 'Data shared with third party? (Y,N,N/A)',
'Headquarter location'];

//User
var authorizedUsers = ['gibson.schnurr@izettle.com', 'linn.andersson@izettle.com',
'josefin.eklund@izettle.com', 'maaike.gerritse@izettle.com', 'markus.kanerva@izettle.com',
'roxanne.baumann@izettle.com', 'shumel.rahman@izettle.com'];
var currentUser = Session.getActiveUser().getEmail();
var numberOfAuthUsers = (authorizedUsers.length - 1);
