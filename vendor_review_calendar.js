function event_creation_validation(existingEventDbOnedArray, appID) {
  for (c = 0; c < existingEventDbOnedArray.length; c++) {
    var dbRow = c + 2;
    if (existingEventDbOnedArray[c] == appID) {
      logs_tst('An event already exists for an application with this ID ' + appID);
      var eventExists = 'yes';
      return eventExists;
    }
    else if (existingEventDbOnedArray[c] != appID && c != (existingEventDbOnedArray.length - 1)) {
      continue;
    }
    else if (existingEventDbOnedArray[c] != appID && c == (existingEventDbOnedArray.length - 1)) {
      logs_tst('Event for ID does not exist ' + appID);
      var eventExists = 'no';
      return eventExists;
    }
  }
}

//remove notice period and create a review date event 180 days ahead of contract expiration date include two events in the calendar for each application
//if the current date is > date at wich the calendar event would be created. do not create
function create_single_day_events(qOneReviewStartDate, qTwoReviewStartDate, qThreeReviewStartDate, qFourReviewStartDate, agreeEndDate, appID, appName, appMan) {

  var reviewCalendar = CalendarApp.getCalendarById('izettle.com_d7p21j601qoq1rih87qnhch9lc@group.calendar.google.com');
//I have started to create the events would still need the if function to determine if this would be appropriate way to do This
//if it is then we could do a date evaluation to name the events based on quarters with an equation this is also where we would add the if statement based on date of the date run to determine if it should be run or not
  var qOneReviewEvent = reviewCalendar.createAllDayEvent('Expiry | ' + appName + ' | ' + appMan, new Date(agreeEndDate));
  contractEndDateEvent.setDescription('The contract expires on this date.');

  var qTwoReviewEvent = reviewCalendar.createAllDayEvent('Expiry | ' + appName + ' | ' + appMan, new Date(agreeEndDate));
  contractEndDateEvent.setDescription('The contract expires on this date.');

  var qThreeReviewEvent = reviewCalendar.createAllDayEvent('Expiry | ' + appName + ' | ' + appMan, new Date(agreeEndDate));
  contractEndDateEvent.setDescription('The contract expires on this date.');

  var qFourReviewEvent = reviewCalendar.createAllDayEvent('Expiry | ' + appName + ' | ' + appMan, new Date(agreeEndDate));
  contractEndDateEvent.setDescription('The contract expires on this date.');

  var contractEndDateEvent = reviewCalendar.createAllDayEvent('Expiry | ' + appName + ' | ' + appMan, new Date(agreeEndDate));
  contractEndDateEvent.setDescription('The contract expires on this date.');

  logs_tst('Begin Customization');

  if (appMan == 'Shumel Rahman') {
    lastNoticeDayEvent.setColor(3);
    contractEndDateEvent.setColor(3);
  }
  else if (appMan == 'Maaike Gerritse') {
    lastNoticeDayEvent.setColor(4);
    contractEndDateEvent.setColor(4);
  }
  else if (appMan == 'Josefin Eklund') {
    lastNoticeDayEvent.setColor(6);
    contractEndDateEvent.setColor(6);
  }
  else if (appMan == 'Roxanne Baumann') {
    lastNoticeDayEvent.setColor(10);
    contractEndDateEvent.setColor(10);
  }
  else if (appMan == 'Markus Kanerva') {
    lastNoticeDayEvent.setColor(1);
    contractEndDateEvent.setColor(1);
  }
  else if (appMan == 'Gibson Schnurr') {
    lastNoticeDayEvent.setColor(7);
    contractEndDateEvent.setColor(7);
  }
  else {
    lastNoticeDayEvent.setColor(8);
    contractEndDateEvent.setColor(8);
  }
  logs_tst('EVENTS CREATED');

}

//if the current date is > date at wich the calendar event would be created. do not create
function create_renewal_calendar() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var rcdb = ss.getSheetByName('Renewal_Calendar_DB');
  var rcdbOrigLr = rcdb.getLastRow();
  var rcdbOrigLc = rcdb.getLastColumn();
  var rcdbTitleColumnArr = rcdb.getRange(1, 1, 1, rcdbOrigLc).getValues();
  var rcdbTcaOned = flatten_arr(rcdbTitleColumnArr);
  var rcdbIdColPos = find_col(rcdbTcaOned, 'SL-ID');

  var tcaOned = flatten_arr(tpslTitleColumnArr);
  var tcaAppIdColPos = find_col(tcaOned, 'SL-ID');
  var tcaAppColPos = find_col(tcaOned, 'Application');
  var tcaVendClassColPos = find_col(tcaOned, 'Vendor/Supplier Classification');
  var tcaAppManColPos = find_col(tcaOned, 'Application Manager');
  var tcaBusOwnColPos = find_col(tcaOned, 'Business System Owner');
  var tcaAgreeEndDateColPos = find_col(tcaOned, 'Agreement End Date');
  var tcaAgreeStartDateColPos = find_col(tcaOned, 'Initial Agreement Start Date');
  logs_tst('All ColPos Found');

  var missingAppInfoArr = [];


  var applicationArr = tpsl.getRange(4, tcaAppColPos, tpslLr, 1).getValues();
  var appArrOned = flatten_arr(applicationArr);

  for (var a = 0; a < appArrOned.length; a++) {
    // for loop that loops through the app array
    var appRow = a + 4;
    logs_tst('Application Row = ' + appRow);
    var vendClass = tpsl.getRange(appRow, tcaVendClassColPos, 1, 1).getValue();
    logs_tst('Vendor Class = ' + vendClass);
    // basic information found such as app row and vendor class
    if (vendClass == 'Tactical' || vendClass == 'Strategic') {
      logs_tst('____DB VALIDATION BEGUN_____');
      var appID = tpsl.getRange(appRow, tcaAppIdColPos, 1, 1).getValue();
      logs_tst('Application ID = ' + appID);
      var appName = tpsl.getRange(appRow, tcaAppColPos, 1, 1).getValue();
      logs_tst('Application = ' + appName);
      var appMan = tpsl.getRange(appRow, tcaAppManColPos, 1, 1).getValue();
      logs_tst('Application Manager = ' + appMan);
      var agreeEndDate = tpsl.getRange(appRow, tcaAgreeEndDateColPos, 1, 1).getValue();
      if (agreeEndDate == '') {
        missingAppInfoArr.push(appName);
        logs_tst('The application ' + appName + ' is missing an agreement end date.')
        continue;
      }
      else {
        logs_tst('____EVENT CREATION STARTED_____');
        logs_tst('Agreement End Date = ' + agreeEndDate);
        var qOneReviewStartDate = agreeEndDate + (msPerDay * 90);
        var qTwoReviewStartDate = agreeEndDate + (msPerDay * 180);
        var qThreeReviewStartDate = agreeEndDate + (msPerDay * 270);
        var qFourReviewStartDate = agreeEndDate + (msPerDay * 360);
        create_single_day_events(qOneReviewStartDate, qTwoReviewStartDate, qThreeReviewStartDate, qFourReviewStartDate, agreeEndDate, appID, appName, appMan);
      }
    }
    else {
      continue;
    }
  }//calendar loop closed
  var compLogs1 = Logger.getLog();
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'Event Creation Comp Log', compLogs1);
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'Vendor Review Calendar Applications with Missing Information', missingAppInfoArr);
} //end of function

//To make this a function to call in other functions I only need to delete the for loop this should check if it has a y in the edits column
function del_cal_event_from_rcdb() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var reviewCalendar = CalendarApp.getCalendarById('izettle.com_d7p21j601qoq1rih87qnhch9lc@group.calendar.google.com');

  var rcdb = ss.getSheetByName('Renewal_Calendar_DB');
  var rcdbOrigLr = rcdb.getLastRow();
  var rcdbOrigLc = rcdb.getLastColumn();
  var rcdbTitleColumnArr = rcdb.getRange(1, 1, 1, rcdbOrigLc).getValues();
  var rcdbTcaOned = flatten_arr(rcdbTitleColumnArr);
  var rcdbIdColPos = find_col(rcdbTcaOned, 'SL-ID');
  var rcdbAppColPos = find_col(rcdbTcaOned, 'Application');
  var rcdbNotPerEvIdColPos = find_col(rcdbTcaOned, 'Notice Period Event');
  var rcdbLasNotDayEvIdColPos = find_col(rcdbTcaOned, 'Last Notice Day Event');
  var rcdbContEndDateEvIdColPos = find_col(rcdbTcaOned, 'Contract End Date');
  var rcdbCreationDateColPos = find_col(rcdbTcaOned, 'Creation Date');
  var rcdbCommitEditsColPos = find_col(rcdbTcaOned, 'Commit Edits?');

  var rcdbIdArr = rcdb.getRange(2, rcdbIdColPos, rcdbOrigLr, 1).getValues();
  var rcdbIdArrOned = flatten_arr(rcdbIdArr);

  var rcdbIdArrRowAdj = 2;

  for (var dd = 0; dd < rcdbIdArrOned.length; dd++) {
    var rcdbIdRow = dd + rcdbIdArrRowAdj;
    var rcdbCommitEditValue = rcdb.getRange(rcdbIdRow, rcdbCommitEditsColPos, 1, 1).getValue();

    if (rcdbCommitEditValue == 'Y' || rcdbCommitEditValue == 'y' || rcdbCommitEditValue == 'yes' || rcdbCommitEditValue == 'Yes') {
      logs_tst('Commit Edits column indicates that this event should be deleted.');
      var npeId = rcdb.getRange(rcdbIdRow, rcdbNotPerEvIdColPos, 1, 1).getValue();
      var lndeId = rcdb.getRange(rcdbIdRow, rcdbLasNotDayEvIdColPos, 1, 1).getValue();
      var cedId = rcdb.getRange(rcdbIdRow, rcdbContEndDateEvIdColPos, 1, 1).getValue();
      reviewCalendar.getEventById(npeId).deleteEvent();
      reviewCalendar.getEventById(lndeId).deleteEvent();
      reviewCalendar.getEventById(cedId).deleteEvent();
      logs_tst('All events have been deleted for this application');
      rcdb.deleteRow(rcdbIdRow);
      --rcdbIdArrRowAdj;
    }
    else {
      logs_tst('Commit Edits column indicates that this event should NOT be deleted.');
      continue;
    }
  }
}
