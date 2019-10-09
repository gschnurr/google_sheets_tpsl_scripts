/*
Create a script that pulls in the usage and last usage month by user
set a threshold for when we might want to turn off someones use of Scrive (not using in the last 3 months requires a reach out)
create a form that is sent out to each user that asks why they use scrive and for what team.
create a script that will grab all of the answers to the forms and based on form title and format the answers in spreadsheet
Script to check what surveys have been responded to
script to send reminder
**/


function get_usage() {
  var csu = ss.getSheetByName('Current Scrive Users');
  var csuLr = csu.getLastRow();
  var csuLc = csu.getLastColumn();
  var cau = ss.getSheetByName('Current Active Users');
  var cauLr = cau.getLastRow();
  var cauLc = cau.getLastColumn();
  var fqs = ss.getSheetByName('Form Questions');

  var csuTcaArr = csu.getRange(1, 1, 1, csuLc).getValues();
  var csuTcaOned = flatten_arr(csuTcaArr);
  var csuUserColPos = find_col(csuTcaOned, 'User');
  var csuUserArr = csu.getRange(2, csuUserColPos, csuLr, 1).getValues();
  var csuUserOned = flatten_arr(csuUserArr);

  var cauTcaArr = cau.getRange(1, 1, 1, cauLc).getValues();
  var cauTcaOned = flatten_arr(cauTcaArr);
  var cauUserColPos = find_col(cauTcaOned, 'User');
  var cauUserArr = cau.getRange(2, cauUserColPos, cauLr, 1).getValues();
  var cauUserOned = flatten_arr(cauUserArr);

}
