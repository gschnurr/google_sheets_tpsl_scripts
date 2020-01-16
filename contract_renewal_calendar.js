function create_events() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var rcdb = ss.getSheetByName('REnewal_Calendar_DB');
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

  var tcaOned = flatten_arr(tpslTitleColumnArr);
  var tcaAppIdColPos = find_col(tcaOned, 'SL-ID');
  var tcaAppColPos = find_col(tcaOned, 'Application');
  var tcaVendClassColPos = find_col(tcaOned, 'Vendor Classification');
  var tcaAppManColPos = find_col(tcaOned, 'Application Manager');
  var tcaBusOwnColPos = find_col(tcaOned, 'Business System Owner');
  var tcaAgreeEndDateColPos = find_col(tcaOned, 'Agreement End Date');
  var tcaLastNoticePeriodColPos = find_col(tcaOned, 'Last Notice Period');
  var tcaAgreeStartDateColPos = find_col(tcaOned, 'Initial Agreement Start Date');
  logs_tst('All ColPos Found');

  var applicationArr = tpsl.getRange(4, tcaAppColPos, tpslLr, 1).getValues();
  var appArrOned = flatten_arr(applicationArr);
  var renewalCalendar = CalendarApp.getCalendarById('izettle.com_r8m408f1j9rilitkkva4gaaqd8@group.calendar.google.com');

  for (var a = 0; a < appArrOned.length; a++) {
    var appRow = a + 4;
    logs_tst('Application Row = ' + appRow);
    var vendClass = tpsl.getRange(appRow, tcaVendClassColPos, 1, 1).getValue();
    logs_tst('Vendor Class = ' + vendClass);
    var rcdbLr = rcdb.getLastRow();
    var rcdbFirstEmptyRow = rcdbLr + 1;
    var rcdbIdArr = rcdb.getRange(2, rcdbIdColPos, rcdbLr, 1).getValues();
    var rcdbIdArrOned = flatten_arr(rcdbIdArr);
    //insert for loop here to check against database if it has already created dates for this year (check if ID exists in DB if yes then check if date is within year limit)
    if (vendClass == 'Tactical' || vendClass == 'Strategic') {
      logs_tst('____EVENT CREATION STARTED_____');
      var appID = tpsl.getRange(appRow, tcaAppIdColPos, 1, 1).getValue();
      logs_tst('Application ID = ' + appID);
      var appName = tpsl.getRange(appRow, tcaAppColPos, 1, 1).getValue();
      logs_tst('Application = ' + appName);
      var appMan = tpsl.getRange(appRow, tcaAppManColPos, 1, 1).getValue();
      logs_tst('Application Manager = ' + appMan);
      //date
      //Below each of these we will check each of these columns and determine next actions if the data does not exist
      var agreeEndDate = tpsl.getRange(appRow, tcaAgreeEndDateColPos, 1, 1).getValue();
      //if it does not exist add application to specific AM array to email them and skip the application likely break from this loop
      logs_tst('Agreement End Date = ' + agreeEndDate);
      var noticePeriod = tpsl.getRange(appRow, tcaLastNoticePeriodColPos, 1, 1).getValue();
      // already take care of default is 60
      logs_tst('Notice Period = ' + noticePeriod);
      //date
      var lastNoticeDate = agreeEndDate - (msPerDay * noticePeriod);
      // if does not exist we should assume 90 days before contract expiry 
      logs_tst('Last Notice Date = ' + lastNoticeDate);
      var lastStartPeriod = tpsl.getRange(appRow, tcaAgreeStartDateColPos, 1, 1).getValue();
      logs_tst('Original Agreement Start Date = ' + lastStartPeriod);
      if (noticePeriod == '') {
        //date
        var renewStartDate = agreeEndDate - (msPerDay * 60);
        logs_tst('The tpsl notice period data does not exist. The Renewal Start Date will default to ' + renewStartDate);
      }
      else {
        //date
        var renewStartDate = agreeEndDate - (msPerDay * (noticePeriod + 60));
        logs_tst('Notice period exits. The renewal start date will be set at ' + renewStartDate);
      }
      var noticePeriodEvent = renewalCalendar.createAllDayEvent('Two Month Notice Period | ' + appName + ' | ' + appMan, new Date(renewStartDate));
      var lastNoticeDayEvent = renewalCalendar.createAllDayEvent('Notice Period Ends | ' + appName + ' | ' + appMan, new Date(lastNoticeDate));
      var contractEndDateEvent = renewalCalendar.createAllDayEvent('Contract Expiry | ' + appName + ' | ' + appMan, new Date(agreeEndDate));
      var npeId = noticePeriodEvent.getId();
      var lndeId = lastNoticeDayEvent.getId();
      var cedId = contractEndDateEvent.getId();

      // add BO as guest

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

      logs_tst('____DATABASE BACKUP STARTED_____');
      rcdb.getRange(rcdbFirstEmptyRow, rcdbIdColPos, 1, 1).setValue(appID);
      rcdb.getRange(rcdbFirstEmptyRow, rcdbAppColPos, 1, 1).setValue(appName);
      rcdb.getRange(rcdbFirstEmptyRow, rcdbNotPerEvIdColPos, 1, 1).setValue(npeId);
      rcdb.getRange(rcdbFirstEmptyRow, rcdbLasNotDayEvIdColPos, 1, 1).setValue(lndeId);
      rcdb.getRange(rcdbFirstEmptyRow, rcdbContEndDateEvIdColPos, 1, 1).setValue(cedId);
      rcdb.getRange(rcdbFirstEmptyRow, rcdbCreationDateColPos, 1, 1).setValue(date);
      logs_tst('DATABASE UPDATED');

      var compLogs1 = Logger.getLog();
      MailApp.sendEmail('gibson.schnurr@izettle.com', 'Event Creation Comp Log', compLogs1);
      return;
    }
    else {
      continue;
    }
    //if statement end
  }
  //for loop closed
}


// I still need to add in a what if the data does not exist situation what do we do then