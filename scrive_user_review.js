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

function user_form() {
  var tz = ss.getSpreadsheetTimeZone();
  var date = Utilities.formatDate(new Date(), tz, 'MM');
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

  var fqs = ss.getSheetByName('Form Questions');
  var fqsLr = fqs.getLastRow();
  var fqsLc = fqs.getLastColumn();
  var fqsTcaArr = fqs.getRange(1, 1, 1, fqsLc).getValues();
  var fqsTcaOned = flatten_arr(fqsTcaArr);
  var fqsQColPos = find_col(fqsTcaOned, 'Question');
  var fqsQArr = fqs.getRange(2, fqsQColPos, fqsLr, 1).getValues();
  var fqsQOned = flatten_arr(fqsQArr);
  var fqsQTypeColPos = find_col(fqsTcaOned, 'Question Type');
  var fqsQTypeArr = fqs.getRange(2, fqsQTypeColPos, fqsLr, 1).getValues();
  var fqsQTypeOned = flatten_arr(fqsQTypeArr);
  var fqsAnsOptBgColPos = find_col(fqsTcaOned, 'Answer Option 1');

  csuUserOned.pop();
  logs_tst('User Array = ' + csuUserOned);
  logs_tst('Question array = ' + fqsQOned);

  for (var fc = 0; fc < csuUserOned.length; fc++) {
    var csuCurUserRow = fc + 2;
    var curUserTeam = csu.getRange(csuCurUserRow, csuTeamColPos, 1, 1).getValue();
    var curUser = csuUserOned[fc];
    var curUserLastDate = csu.getRange(csuCurUserRow, csuLastUsageDateColPos, 1, 1).getValue();
    var curUserUsage = csu.getRange(csuCurUserRow, csuUsageColPos, 1, 1).getValue();
    var curUserEmail = csu.getRange(csuCurUserRow, csuEmailColPos, 1, 1).getValue();
    logs_tst('The current user is ' + curUser + ' this user is in row ' + csuCurUserRow +
    ' and in team ' + curUserTeam + '. This users usage and last usage date are ' + curUserUsage + ' ' + curUserLastDate +
    '. This users email is ' + curUserEmail);
    if (curUserTeam != 'IT' && curUserTeam != 'Legal') {
      var csuUserForm = FormApp.create('Scrive User Review - ' + curUser + '.').setCollectEmail(true);
      csuUserForm.setTitle('Scrive User Review - ' + curUser + '.');
        var sectionHeader = csuUserForm.addSectionHeaderItem();
          sectionHeader.setTitle('Please Answer the below questions in terms of how you use Scrive to complete your work.');
          if (curUserUsage != '') {
            sectionHeader.setHelpText('Our records indicate that you have recently used scrive to send/complete/sign a document.' +
            'You have completed ' + curUserUsage + ' actions in Scrive in the last 6 months.')
          }
          else if (curUserUsage == '') {
            sectionHeader.setHelpText('Our records indicate that you have NOT used scrive to send/complete/sign a document in the last 6 months.' +
            'You have completed ' + curUserUsage + ' actions in Scrive in the last 6 months.' + 'Please describe any work you do in Scrive below.')
          }
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
          var fqsAnsOptArr = fqs.getRange(curQRow, fqsAnsOptBgColPos, 1, fqsLc).getValues();
          var fqsAnsOptOned = flatten_arr(fqsAnsOptArr);
          fqsAnsOptOned.pop();
          var mcOptions = [];
          var multiChoice = csuUserForm.addMultipleChoiceItem().setRequired(false);
          for (var mc = 0; mc < fqsAnsOptOned.length; mc++) {
            if (fqsAnsOptOned[mc] == '') {
              continue;
            }
            else {
              var vcChoice = multiChoice.createChoice(fqsAnsOptOned[mc]);
              mcOptions.push(vcChoice);
            }
          }
          multiChoice.setTitle(fqsQOned[cq])
            .setChoices(mcOptions);
          logs_tst('Question with the type Mulitple choice has been created with choices '+
          mcOptions + '.');
        }
        else if (fqsQTypeOned[cq] == 'paragraph') {
          var paraItem = csuUserForm.addParagraphTextItem().setRequired(false);
            paraItem.setTitle(fqsQOned[cq])
          logs_tst('Question with the type paragraph has been created.');
        }
        else if (fqsQTypeOned[cq] == '') {
          continue;
        }
        else {
          ui.alert('The question type ' + fqsQTypeOned[cq] + ' is not currently supported please contact Gibson.');
          logs_tst('The question type ' + fqsQTypeOned[cq] + ' is not currently supported');
          continue;
        }
      }
      logs_tst('Form for user ' + curUser + ' created.');
      var responseUrl = csuUserForm.getPublishedUrl();
      var emailSubject = 'Scrive User Review: Please complete in one week';
      var emailBody = {}
      emailBody.htmlBody = 'Hello ' + curUser + ',' + '<br />' + ' <br />' +
      'Please fill out the form in the link provided '+ '<a href=\"' + responseUrl + '">here. </a>' +
      'IT is working to review the current users in Scrive to get a better understanding of the work being performed in the system ' +
      'and to free up inactive license users. The deadline for completion is one week from today.' + '<br />' + 'All the best,' + '<br />' + 'Gibson'
      MailApp.sendEmail(curUserEmail, emailSubject, '', emailBody);
      logs_tst('Email Sent');
      var compLogs = Logger.getLog();
      MailApp.sendEmail('gibson.schnurr@izettle.com', 'Scrive Form Script: Form Sent', compLogs);
    }
    else {
      logs_tst('This user is in the IT or Legal team and therefore no form will be created.');
      continue;
    }
  }
}

//loop through all of the forms and create an arr of the reponses and take the last response
//loop through the responses and baseed on question update the excel sheet
function get_form_resp() {
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
  var csuEmailArr = csu.getRange(2, csuEmailColPos, csuLr, 1).getValues();
  var csuEmailOned = flatten_arr(csuEmailArr);

  var fqs = ss.getSheetByName('Form Questions');
  var fqsLr = fqs.getLastRow();
  var fqsLc = fqs.getLastColumn();
  var fqsTcaArr = fqs.getRange(1, 1, 1, fqsLc).getValues();
  var fqsTcaOned = flatten_arr(fqsTcaArr);
  var fqsQColPos = find_col(fqsTcaOned, 'Question');
  var fqsQArr = fqs.getRange(2, fqsQColPos, fqsLr, 1).getValues();
  var fqsQOned = flatten_arr(fqsQArr);
  var fqsQTypeColPos = find_col(fqsTcaOned, 'Question Type');
  var fqsQTypeArr = fqs.getRange(2, fqsQTypeColPos, fqsLr, 1).getValues();
  var fqsQTypeOned = flatten_arr(fqsQTypeArr);
  var fqsAnsOptBgColPos = find_col(fqsTcaOned, 'Answer Option 1');
  var fqsAnsColHeadColPos = find_col(fqsTcaOned, 'Answer Col Header');

  var formIdArr = [];
  var scriveForms = DriveApp.searchFiles('title contains "Scrive User Review" and mimeType contains "form"');
  while (scriveForms.hasNext()) {
    var curFormIt = scriveForms.next();
    formIdArr.push(curFormIt.getId());
  }

  for (var fi = 0; fi < formIdArr.length; fi++) {
    var curForm = FormApp.openById(formIdArr[fi]);
    logs_tst('The Form iD ' + formIdArr[fi]);
    var curFormItemArr = curForm.getItems();
    var curFormResponses = curForm.getResponses();
    var formLastResponseInt = curFormResponses.length - 1;
    var curFormLatestResp = curFormResponses[formLastResponseInt];
    var curFormUserEmail = curFormLatestResp.getRespondentEmail();
    logs_tst('The respondent to this form was ' + curFormUserEmail);
    for (var fe = 0; fe < csuEmailOned.length; fe++){
      if (csuEmailOned[fe] == curFormUserEmail) {
        var curFormUserRow = fe + 2;
        logs_tst('User Email found in CSU, user is in row ' + curFormUserRow);
        break;
      }
      else if (csuEmailOned[fe] != curFormUserEmail && fe != (csuEmailOned.length - 1)) {
        logs_tst('User not found yet be we are not at the end of the array.');
        continue;
      }
      else if (csuEmailOned[fe] != curFormUserEmail && fe == (csuEmailOned.length - 1)) {
        var curFormUserRow = '';
        logs_tst('User not found user row is blank.');
        break;
      }
      else {
        logs_tst('something went wrong #1');
        ui.alert('something went wrong #1');
      }
    }
    if (curFormUserRow == '') {
      logs_tst('User email ' + curFormUserEmail + ' does not exist in the current scrive users sheet.');
      break;
    }
    for (var ia = 0; ia < curFormItemArr.length; ia++) {
      var curItemType = curFormItemArr[ia].getType();
      if (curItemType == FormApp.ItemType.TEXT || curItemType == FormApp.ItemType.MULTIPLE_CHOICE || curItemType == FormApp.ItemType.PARAGRAPH_TEXT) {
        var itemQuest = curFormItemArr[ia].getTitle();
        logs_tst('The question is ' + itemQuest);
        var curItemId = curFormItemArr[ia].getId();
        logs_tst('The item ID is ' + curItemId);
        var questIndex = fqsQOned.indexOf(itemQuest);
        logs_tst('The index of this question in the question array is ' + questIndex);
        var questRow = questIndex + 2;
        var answerCol = fqs.getRange(questRow, fqsAnsColHeadColPos, 1, 1).getValue();
        var answerColPosCsu = find_col(csuTcaOned, answerCol);
        var itemResponseArr = curFormLatestResp.getItemResponses();
        for (var er = 0; er < itemResponseArr.length; er++) {
          var curItemResp = itemResponseArr[er];
          var curItemRespId = curItemResp.getItem().getId();
          if (curItemRespId == curItemId) {
            var userAnswer = curItemResp.getResponse();
            logs_tst('The coordinates of the answer cell are (row, col)' + '(' + curFormUserRow + ',' + answerColPosCsu + ')');
            break;
          }
          else if (curItemRespId != curItemId && er != (itemResponseArr.length -1)) {
            logs_tst('ID not found but we are not at the end of the array yet.')
            continue;
          }
          else if (curItemRespId != curItemId && er == (itemResponseArr.length -1)) {
            var userAnswer = csu.getRange(curFormUserRow, answerColPosCsu, 1, 1).getValue();
          }
        }
        logs_tst('The answer for Question ' + itemQuest + ' is ' + userAnswer);
        csu.getRange(curFormUserRow, answerColPosCsu, 1, 1).setValue(userAnswer);
        logs_tst('Answer set in CSU sheet');
      }
      else {
        continue;
      }
    }
  }
  logs_tst(formIdArr);
  var compLogs = Logger.getLog();
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'Scrive Form Script: Get Forms', compLogs);
}

//this function will loop through all of the forms and see how many have responses
//if a form does not have a response it will push the form name into an array which will be presented as a list in my emailTo

function check_for_resp() {

  var formIdArr = [];
  var scriveForms = DriveApp.searchFiles('title contains "Scrive User Review" and mimeType contains "form"');
  while (scriveForms.hasNext()) {
    var curFormIt = scriveForms.next();
    formIdArr.push(curFormIt.getId());
  }

  for (var cfr = 0; cfr < formIdArr.length; cfr++) {
    var curForm = FormApp.openById(formIdArr[cfr]);
    var formName = curForm.getTitle();
    var formResponses = curForm.getResponses();
    var fromLatestResp = formResponses.length - 1;
    var mostRecentResp = formResponses[fromLatestResp];
    if (fromLatestResp < 0) {
      var respondingUser = 'There are no responses';
    }
    else {
      var respondingUser = mostRecentResp.getRespondentEmail();
    }
    logs_tst('Form Title = ' + formName + '. This form has ' + formResponses.length + '. The latest reponse is from ' + respondingUser);
  }
  logs_tst(formIdArr);
  var compLogs = Logger.getLog();
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'Scrive Form Script: check for responses', compLogs);
}
