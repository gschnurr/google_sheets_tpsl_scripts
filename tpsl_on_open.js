function onOpen(e) {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var spreadsheetName = spreadsheet.getName();
  var sheetName = ss.getSheetName();

  ui.createMenu('AM Tools')
    .addItem('Add Application', 'gen_export_ins')
    .addSeparator()
    .addItem('Archive Application', 'gen_export_ins')
    .addSeparator()
    .addSubMenu(ui.createMenu('Generic Export')
      .addItem('Generic Export Instructions', 'gen_export_ins')
      .addSeparator()
      .addItem('Generate Export', 'gen_export_wrapper')
      .addSeparator()
      .addItem('Clean Export', 'clean_export_wrapper'))
    .addSeparator()
    .addSubMenu(ui.createMenu('PP Exports')
      .addItem('PP Quarterly Macro Instructions', 'pp_information_prompt')
      .addSeparator()
      .addItem('PayPal GDPR Extract', 'tpsl_pp_extract')
      .addSeparator()
      .addItem('Send Forms', 'pp_bo_form')
      .addSeparator()
      .addItem('Get Updates', 'get_updates')
      .addSeparator()
      .addItem('Push Updates', 'pp_push_updates_wrapper'))
    .addToUi();

  //Checking if clean export should be run which is defined first by spreadsheet name then by number sheets
  if (spreadsheetName == exportName && sheets.length > 2) {
    clean_export_wrapper();
  }
};
