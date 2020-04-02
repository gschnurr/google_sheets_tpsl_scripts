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

function create_single_day_events(noticePeriod, agreeEndDate, appID, appName, appMan, rcdbFirstEmptyRow) {

  var renewalCalendar = CalendarApp.getCalendarById('izettle.com_d7p21j601qoq1rih87qnhch9lc@group.calendar.google.com');

  var rcdb = ss.getSheetByName('Renewal_Calendar_DB');
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

  if (noticePeriod == '' || noticePeriod == 0) {
    var lastNoticeDate = agreeEndDate - (msPerDay * 90);
    var renewStartDate = lastNoticeDate - (msPerDay * 60);
    logs_tst('The tpsl notice period data does not exist. The Renewal Start Date will default to ' + renewStartDate);
  }
  else {
    var lastNoticeDate = agreeEndDate - (msPerDay * noticePeriod);
    var renewStartDate = lastNoticeDate - (msPerDay * 60);
    logs_tst('Notice period exits. The renewal start date will be set at ' + renewStartDate);
  }

  var noticePeriodEvent = renewalCalendar.createAllDayEvent('60 Day | ' + appName + ' | ' + appMan, new Date(renewStartDate));
  noticePeriodEvent.setDescription('This is a reminder 60 days before prior written notice of termination is due to the vendor. ' +
  'The application manager should begin to work with the business owner to renew or terminate the contract for this vendor.');

  var lastNoticeDayEvent = renewalCalendar.createAllDayEvent('Notice | ' + appName + ' | ' + appMan, new Date(lastNoticeDate));
  lastNoticeDayEvent.setDescription('This the last day that prior written notice of termination can be submitted to the vendor.');

  var contractEndDateEvent = renewalCalendar.createAllDayEvent('Expiry | ' + appName + ' | ' + appMan, new Date(agreeEndDate));
  contractEndDateEvent.setDescription('The contract expires on this date.');

  logs_tst('Begin Customization');

  if (appMan == 'Shumel Rahman') {
    noticePeriodEvent.setColor(3);
    lastNoticeDayEvent.setColor(3);
    contractEndDateEvent.setColor(3);
  }
  else if (appMan == 'Maaike Gerritse') {
    noticePeriodEvent.setColor(4);
    lastNoticeDayEvent.setColor(4);
    contractEndDateEvent.setColor(4);
  }
  else if (appMan == 'Josefin Eklund') {
    noticePeriodEvent.setColor(6);
    lastNoticeDayEvent.setColor(6);
    contractEndDateEvent.setColor(6);
  }
  else if (appMan == 'Roxanne Baumann') {
    noticePeriodEvent.setColor(10);
    lastNoticeDayEvent.setColor(10);
    contractEndDateEvent.setColor(10);
  }
  else if (appMan == 'Markus Kanerva') {
    noticePeriodEvent.setColor(1);
    lastNoticeDayEvent.setColor(1);
    contractEndDateEvent.setColor(1);
  }
  else if (appMan == 'Gibson Schnurr') {
    noticePeriodEvent.setColor(7);
    lastNoticeDayEvent.setColor(7);
    contractEndDateEvent.setColor(7);
  }
  else {
    noticePeriodEvent.setColor(8);
    lastNoticeDayEvent.setColor(8);
    contractEndDateEvent.setColor(8);
  }
  logs_tst('EVENTS CREATED');

  var npeId = noticePeriodEvent.getId();
  var lndeId = lastNoticeDayEvent.getId();
  var cedId = contractEndDateEvent.getId();

  logs_tst('____DATABASE BACKUP STARTED_____');
  rcdb.getRange(rcdbFirstEmptyRow, rcdbIdColPos, 1, 1).setValue(appID);
  rcdb.getRange(rcdbFirstEmptyRow, rcdbAppColPos, 1, 1).setValue(appName);
  rcdb.getRange(rcdbFirstEmptyRow, rcdbNotPerEvIdColPos, 1, 1).setValue(npeId);
  rcdb.getRange(rcdbFirstEmptyRow, rcdbLasNotDayEvIdColPos, 1, 1).setValue(lndeId);
  rcdb.getRange(rcdbFirstEmptyRow, rcdbContEndDateEvIdColPos, 1, 1).setValue(cedId);
  rcdb.getRange(rcdbFirstEmptyRow, rcdbCreationDateColPos, 1, 1).setValue(date);
  logs_tst('DATABASE UPDATED');
}

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
  var tcaLastNoticePeriodColPos = find_col(tcaOned, 'Last Notice Period');
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
    var rcdbLr = rcdb.getLastRow();
    var rcdbFirstEmptyRow = rcdbLr + 1;
    var rcdbIdArr = rcdb.getRange(2, rcdbIdColPos, rcdbLr, 1).getValues();
    var rcdbIdArrOned = flatten_arr(rcdbIdArr);
    // basic information found such as app row and vendor class
    if (vendClass == 'Tactical' || vendClass == 'Strategic' || vendClass == 'Operational' || vendClass == 'Commodity') {
      logs_tst('____DB VALIDATION BEGUN_____');
      var appID = tpsl.getRange(appRow, tcaAppIdColPos, 1, 1).getValue();
      logs_tst('Application ID = ' + appID);
      var appName = tpsl.getRange(appRow, tcaAppColPos, 1, 1).getValue();
      logs_tst('Application = ' + appName);
      var appMan = tpsl.getRange(appRow, tcaAppManColPos, 1, 1).getValue();
      logs_tst('Application Manager = ' + appMan);

      var eventExists = event_creation_validation(rcdbIdArrOned, appID);

      if (eventExists == 'yes') {
        continue;
      }
      else if (eventExists == 'no') {
        var agreeEndDate = tpsl.getRange(appRow, tcaAgreeEndDateColPos, 1, 1).getValue();
        if (agreeEndDate == '') {
          missingAppInfoArr.push(appName);
          logs_tst('The application ' + appName + ' is missing an agreement end date.')
          continue;
        }
        else {
          logs_tst('____EVENT CREATION STARTED_____');
          logs_tst('Agreement End Date = ' + agreeEndDate);
          var noticePeriod = tpsl.getRange(appRow, tcaLastNoticePeriodColPos, 1, 1).getValue();
          logs_tst('Notice Period = ' + noticePeriod);
          var lastStartPeriod = tpsl.getRange(appRow, tcaAgreeStartDateColPos, 1, 1).getValue();
          logs_tst('Original Agreement Start Date = ' + lastStartPeriod);
          create_single_day_events(noticePeriod, agreeEndDate, appID, appName, appMan, rcdbFirstEmptyRow);
        }
      }
    }
    else {
      continue;
    }
  }//calendar loop closed
  var compLogs1 = Logger.getLog();
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'Event Creation Comp Log', compLogs1);
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'Renewal Calendar Applications with Missing Information', missingAppInfoArr);
} //end of function

//To make this a function to call in other functions I only need to delete the for loop this should check if it has a y in the edits column
function del_multiple_cal_event_from_rcdb() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var renewalCalendar = CalendarApp.getCalendarById('izettle.com_d7p21j601qoq1rih87qnhch9lc@group.calendar.google.com');

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
      renewalCalendar.getEventById(npeId).deleteEvent();
      renewalCalendar.getEventById(lndeId).deleteEvent();
      renewalCalendar.getEventById(cedId).deleteEvent();
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

function delete_single_id_event(renewalCalendar, existingEventDbOnedArray, appID, npeColPos, lndeColPos, cedColPos, rcdbSheet) {
  for (c = 0; c < existingEventDbOnedArray.length; c++) {
    var dbRow = c + 2;
    if (existingEventDbOnedArray[c] == appID) {
      logs_tst('Events with the following SL-ID found and will be deleted and replaced: ' + appID);
      var npeId = rcdb.getRange(dbRow, npeColPos, 1, 1).getValue();
      var lndeId = rcdb.getRange(dbRow, lndeColPos, 1, 1).getValue();
      var cedId = rcdb.getRange(dbRow, cedColPos, 1, 1).getValue();
      renewalCalendar.getEventById(npeId).deleteEvent();
      renewalCalendar.getEventById(lndeId).deleteEvent();
      renewalCalendar.getEventById(cedId).deleteEvent();
      rcdbSheet.deleteRow(dbRow);
    }
    else if (existingEventDbOnedArray[c] != appID && c != (existingEventDbOnedArray.length - 1)) {
      continue;
    }
    else if (existingEventDbOnedArray[c] != appID && c == (existingEventDbOnedArray.length - 1)) {
      logs_tst('Events do not exist for appID: ' + appID);
    }
  }
}

function update_event(appIdRow){
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var renewalCalendar = CalendarApp.getCalendarById('izettle.com_d7p21j601qoq1rih87qnhch9lc@group.calendar.google.com');

  var tcaOned = flatten_arr(tpslTitleColumnArr);
  var tcaAppIdColPos = find_col(tcaOned, 'SL-ID');
  var tcaAppColPos = find_col(tcaOned, 'Application');
  var tcaVendClassColPos = find_col(tcaOned, 'Vendor/Supplier Classification');
  var tcaAppManColPos = find_col(tcaOned, 'Application Manager');
  var tcaBusOwnColPos = find_col(tcaOned, 'Business System Owner');
  var tcaAgreeEndDateColPos = find_col(tcaOned, 'Agreement End Date');
  var tcaLastNoticePeriodColPos = find_col(tcaOned, 'Last Notice Period');
  var tcaAgreeStartDateColPos = find_col(tcaOned, 'Initial Agreement Start Date');

  var rcdb = ss.getSheetByName('Renewal_Calendar_DB');
  var rcdbOrigLr = rcdb.getLastRow();
  var rcdbOrigLc = rcdb.getLastColumn();
  var rcdbTitleColumnArr = rcdb.getRange(1, 1, 1, rcdbOrigLc).getValues();
  var rcdbTcaOned = flatten_arr(rcdbTitleColumnArr);
  var rcdbIdColPos = find_col(rcdbTcaOned, 'SL-ID');
  var rcdbNotPerEvIdColPos = find_col(rcdbTcaOned, 'Notice Period Event');
  var rcdbLasNotDayEvIdColPos = find_col(rcdbTcaOned, 'Last Notice Day Event');
  var rcdbContEndDateEvIdColPos = find_col(rcdbTcaOned, 'Contract End Date');
  var rcdbIdArr = rcdb.getRange(2, rcdbIdColPos, rcdbOrigLr, 1).getValues();
  var rcdbIdArrOned = flatten_arr(rcdbIdArr);

  var editAppId = tpsl.getRange(appIdRow, tcaAppIdColPos,1 ,1).getValue();

  var noticePeriod = tpsl.getRange(appIdRow, tcaLastNoticePeriodColPos, 1, 1).getValue();
  var agreeEndDate = tpsl.getRange(appIdRow, tcaAgreeEndDateColPos, 1, 1).getValue();
  var appName = tpsl.getRange(appIdRow, tcaAppColPos, 1, 1).getValue();
  var appMan = tpsl.getRange(appIdRow, tcaAppManColPos, 1, 1).getValue();

  var eventExists = event_creation_validation(rcdbIdArrOned, editAppId);

  if (eventExists == 'yes') {
    delete_single_id_event(renewalCalendar, rcdbIdArrOned, editAppId, rcdbNotPerEvIdColPos, rcdbLasNotDayEvIdColPos, rcdbContEndDateEvIdColPos, rcdb);
    var rcdbLr = rcdb.getLastRow();
    var rcdbFirstEmptyRow = rcdbLr + 1;
    create_single_day_events(noticePeriod, agreeEndDate, editAppId, appName, appMan, rcdbFirstEmptyRow)
  }
  else if (eventExists == 'no') {
    var rcdbLr = rcdb.getLastRow();
    var rcdbFirstEmptyRow = rcdbLr + 1;
    create_single_day_events(noticePeriod, agreeEndDate, editAppId, appName, appMan, rcdbFirstEmptyRow)
  }
  else{
    logs_tst('Something went wrong -- Event exists returned neither yes or no');
  }
}
