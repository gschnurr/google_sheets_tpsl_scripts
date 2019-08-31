/** @OnlyCurrentDoc */


function onOpen(e) {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var exportName = 'Generic Export ' + date;
  var spreadsheetName = spreadsheet.getName();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var ui = SpreadsheetApp.getUi();

  //Macro Menus
    ui.createMenu('Simple Exports')
      .addItem('Generic Export Instructions', 'gen_export_ins')
      .addSeparator()
      .addItem('Generate Export', 'gen_export_wrapper')
      .addSeparator()
      .addItem('Clean Export', 'clean_export_wrapper')
      .addToUi();

    ui.createMenu('PP Exports')
      .addItem('PP Quarterly Macro Instructions', 'pp_information_prompt')
      .addSeparator()
      .addItem('PayPal GDPR Extract', 'tpsl_pp_extract')
      .addSeparator()
      .addItem('Get Updates', 'get_updates')
      .addSeparator()
      .addItem('Push Updates', 'pp_push_updates_wrapper')
      .addToUi();


  //Checking if clean export should be run which is defined first by spreadsheet name then by number sheets
  if (spreadsheetName == exportName && sheets.length > 2) {
    clean_export_wrapper();
  }
};
