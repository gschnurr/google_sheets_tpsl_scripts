/*
Create a script that pulls in the usage and last usage month by user - done
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
  var csuUsageColPos = find_col(csuTcaOned, 'Usage');
  var csuLastUsageDateColPos = find_col(csuTcaOned, 'Last Date Used');
  var csuUserArr = csu.getRange(2, csuUserColPos, csuLr, 1).getValues();
  var csuUserOned = flatten_arr(csuUserArr);

  var cauTcaArr = cau.getRange(1, 1, 1, cauLc).getValues();
  var cauTcaOned = flatten_arr(cauTcaArr);
  var cauUserColPos = find_col(cauTcaOned, 'User');
  var cauClosedDocColPos = find_col(cauTcaOned, 'Closed');
  var cauSentDocColPos = find_col(cauTcaOned, 'Sent');
  var cauSignedDocColPos = find_col(cauTcaOned, 'Signed');
  var cauDateColPos = find_col(cauTcaOned, 'Date');
  var cauUserArr = cau.getRange(2, cauUserColPos, cauLr, 1).getValues();
  var cauUserOned = flatten_arr(cauUserArr);

  for (var su = 0; su < csuUserOned.length; su++) {
    var csuCurRow = su + 2;
    for (var au = 0; au < cauUserOned.length; au++) {
      if (csuUserOned[su] == cauUserOned[au]) {
        var cauCurRow = au + 2;
        var closedDocs = cau.getRange(cauCurRow, cauClosedDocColPos, 1, 1).getValue();
        var sentDocs = cau.getRange(cauCurRow, cauSentDocColPos, 1, 1).getValue();
        var signedDocs = cau.getRange(cauCurRow, cauSignedDocColPos, 1, 1).getValue();
        var userLineUsage = closedDocs + sentDocs + signedDocs;
        var userUsage = csu.getRange(csuCurRow, csuUsageColPos, 1, 1).getValue();
        var userTotalUsage = userLineUsage + userUsage;
        var userLastUsageDate = csu.getRange(csuCurRow, csuLastUsageDateColPos, 1, 1).getValue();

        csu.getRange(csuCurRow, csuUsageColPos, 1, 1).setValue(userTotalUsage);

        if (userLastUsageDate == '') {
          var curMonth = cau.getRange(cauCurRow, cauDateColPos, 1, 1).getValue();
          csu.getRange(csuCurRow, csuLastUsageDateColPos, 1, 1).setValue(curMonth);
        }
        else {
          continue;
        }
      }
      else {
        continue;
      }
    }
  }
}


//set a threshold for when we might want to turn off someones use of Scrive (not using in the last 3 months requires a reach out)
//create a form that is sent out to each user that asks why they use scrive and for what team.
//loop through all of the users who are not IT or Legal
//check for last used date say you have not sent/sign/created a contract in the last x months send to email and use the user name
//Loop through and create forms based on questions send to everyone by the email provided

// I just need to add support for it to send emails to the appropriate user and to send logs to me then I can test
function user_form() {
  var tz = ss.getSpreadsheetTimeZone();
  var date = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy');
  logs_tst('Date establish: ' + date);

  var ui = SpreadsheetApp.getUi();

  var csu = ss.getSheetByName('Current Scrive Users');
  var csuLr = csu.getLastRow();
  var csuLc = csu.getLastColumn();
  var csuTcaArr = csu.getRange(1, 1, 1, csuLc).getValues();
  var csuTcaOned = flatten_arr(csuTcaArr);
  var csuUserColPos = find_col(csuTcaOned, 'User');
  var csuUsageColPos = find_col(csuTcaOned, 'Usage');
  var csuLastUsageDateColPos = find_col(csuTcaOned, 'Last Date Used');
  var csuTeamColPos = find_col(csuTcaOned, 'Team');
  var csuEmailColPos = find_col(csuTcaOned, 'Email');
  var csuUserArr = csu.getRange(2, csuUserColPos, csuLr, 1).getValues();
  var csuUserOned = flatten_arr(csuUserArr);

  var cau = ss.getSheetByName('Current Active Users');
  var cauLr = cau.getLastRow();
  var cauLc = cau.getLastColumn();

  var fqs = ss.getSheetByName('Form Questions');
  var fqsLr = fqs.getLastRow();
  var fqsLc = fqs.getLastColumn();
  var fqsTcaArr = fqs.getRange(1, 1, 1, fqsLc);
  var fqsTcaOned = flatten_arr(fqsTcaArr);
  var fqsQColPos = find_col(fqsTcaOned, 'Question');
  var fqsQArr = fqs.getRange(2, fqsQColPos, fqsLr, 1);
  var fqsQOned = flatten_arr(fqsQArr);
  var fqsQTypeColPos = find_col(fqsTcaOned, 'Question Type');
  var fqsQTypeArr = fqs.getRange(2, fqsQTypeColPos, fqsLr, 1);
  var fqsQTypeOned = flatten_arr(fqsQTypeArr);
  var fqsAnsOptBgColPos = find_col(fqsTcaOned, 'Answer Option 1');

  csuUserOned.pop();
  logs_tst('User Array = ' + csuUserOned);

  for (var fc = 0; fc < csuUserOned.length; fc++) {
    var csuCurUserRow = fc + 2;
    var curUserTeam = csu.getRange(csuCurUserRow, csuTeamColPos, 1, 1).getValue();
    var curUser = csuUserOned[fc];
    logs_tst('The current user is ' + curUser + ' this user is in row ' + csuCurUserRow +
    ' and in team ' + curUserTeam + '.');
    if (curUserTeam != 'IT' || curUserTeam != 'Legal') {
      var csuUserForm = FormApp.create('Scrive User Review - ' + curUser + '.');
      logs_tst(csuUserForm + ' created.');
      for (var cq = 0; cq < fqsQOned.length; cq++) {
        logs_tst('The question is ' + fqsQOned[cq] + '. This question is of the type ' +
        fqsQTypeOned[cq]);
        if (fqsQTypeOned[cq] == 'open') {
          var textItem = csuUserForm.addTextItem().setRequired(false);
            textItem.setTitle(fqsQOned[cq])
          logs_tst('Question with the type open has been created.');
        }
        else if (fqsQTypeOned[cq] == 'mc') {
          var curQRow = cq + 2;
          var mcOptionsArr = fqs.getRange(curQRow, fqsAnsOptBgColPos, 1, fqsLc);
          var mcOptionsOned = flatten_arr(mcOptionsArr);
          mcOptionsOned.pop();
          var multiChoice = csuUserForm.addMultipleChoiceItem().setRequired(false);
            multiChoice.setTitle(fqsQOned[cq]);
              .setChoices(mcOptionsOned);
          logs_tst('Question with the type Mulitple choice has been created with choices '+
          mcOptionsOned + '.');
        }
        else {
          ui.alert('The quetion type ' + fqsQTypeOned[cq] + ' is not currently supported please contact Gibson.');
          logs_tst('The quetion type ' + fqsQTypeOned[cq] + ' is not currently supported');
          continue;
        }
      }
    }
    else {
      logs_tst('This user is in the IT or Legal team and therefore no form will be created.');
      continue;
    }
  }
}
