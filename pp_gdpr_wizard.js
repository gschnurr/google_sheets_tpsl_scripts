function pp_wizard_information_prompt() {
  var htmlOutput = HtmlService
    .createHtmlOutput ('<p>For this review you should review the GDPR information for the applications that you are the assigned business system owner. You can update the information or use the Update Wizard. To use the Wizard see the steps below, the purpose of this information.</p>' +
    '<p> Purpose: Type up the purpose of this update exercise.</p>' +
    '<p> 1) To use the Update Wizard start the wizard by going to the GDPR Update Wizard Menu item and select Run Wizard.</p>' +
    '<p> 2) The wizard will cycle through each application that you are the assigned business owner. You will be presented with a prompt for each column.</p>' +
    '<p> 3) This prompt will contain the Application Name, Column Title, Column Description, and Current Information.</p>' +
    '<p> 4) If the current information is correct press the cancel button.</p>' +
    '<p> 5) If the current information is incorrect type the correct information in the free text field and press OK. The cell will automatically be updated with the new information.</p>' +
    '<p> 7) If you would like to skip to the next application type "next application" in the free text field and press OK.</p>' +
    '<p> 8) If you would like to exit the the wizard completely type "ESCAPE" in the free text field and press OK. This is case sensative.</p>'
  )
    .setTitle('GDPR Update Wizard Instructions');
  ui.showSidebar(htmlOutput);
}

function pp_gdpr_wizard() {
  pp_wizard_information_prompt();

  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ppe = ss.getSheetByName('PayPal Extract');
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  var ppeTitleColumnArr = ppe.getRange(1, 1, 1, ppeLc).getValues();
  var currentWizUser = Session.getActiveUser().getEmail();

  var ppeOned = flatten_arr(ppeTitleColumnArr);

  var recordUpdatesColPos = find_col(ppeOned, 'Updates? (Y/N) If yes please make the updates in this sheet');

  var busSysOwnColPos = find_col(ppeOned, 'Business System Owner');
  var busSysOwnerArr = ppe.getRange(2, busSysOwnColPos, ppeLr, 1).getValues();
  var bsoOned = flatten_arr(busSysOwnerArr);

  var appNameColPos = find_col(ppeOned, 'Application');
  var appNameArr = ppe.getRange(2, appNameColPos, ppeLr, 1).getValues();
  var anOned = flatten_arr(appNameArr);

  //this is an array of the row number of the applications owned by the currentWizUser
  var wizUserAppsRowNumArr = [];
  var wizUserAppsArr = [];
  for (var f = 0; f < bsoOned.length; f++) {
    if (bsoOned[f] == currentWizUser) {
      var cwuAppRow = f + 2;
      wizUserAppsRowNumArr.push(cwuAppRow);
      var appNameCell = ppe.getRange(cwuAppRow, appNameColPos, 1, 1).getValue();
      wizUserAppsArr.push(appNameCell);
    }
    else {
      continue;
    }
  }
  //PPE cols doesnt work
  //creates an arr of all col pos of columns to be updated
  var gdprWizColPosArr = [];
  for (var c = 0; c < gdprWizTColArr.length; c++) {
    var gdprWizColPos = find_col(ppeOned, gdprWizTColArr[c]);
    gdprWizColPosArr.push(gdprWizColPos);
  }

  for (var u = 0; u < wizUserAppsRowNumArr.length; u++) {
    for (var v = 0; v < gdprWizColPosArr.length; v++) {
      var appToUpdate = wizUserAppsArr[u];
      var colTitle = gdprWizTColArr[v];
      var colDescription = ppe.getRange(1, gdprWizColPosArr[v], 1, 1).getNotes();
      ppe.getRange(wizUserAppsRowNumArr[u], appNameColPos, 1, 1).activateAsCurrentCell();
      var currentInfo = ppe.getRange(wizUserAppsRowNumArr[u], gdprWizColPosArr[v], 1, 1).getValue();
      var currentInfoCheck;
      if (currentInfo == '') {
        currentInfoCheck = '<BLANK>';
      }
      else {
        currentInfoCheck = currentInfo;
      }
      var wizUserResp = Browser.inputBox(appToUpdate,
        'Please update the information relating to the ' + colTitle + ' field' +
        '. The description of this field is: ' + colDescription +
        ' The current information is '+ currentInfoCheck +
        '. Type your updates in the response field and press OK to accept the changes. ' +
        'If you do not have any changes to make press cancel to move to the next field for review. ',
        Browser.Buttons.OK_CANCEL);
      if (wizUserResp == 'cancel') {
        continue;
      }
      else if (wizUserResp == '') {
        continue;
      }
      else if (wizUserResp == 'next application') {
        break;
      }
      else if (wizUserResp == 'ESCAPE') {
        return;
      }
      else {
        ppe.getRange(wizUserAppsRowNumArr[u], gdprWizColPosArr[v], 1, 1).setValue(wizUserResp);
        ppe.getRange(wizUserAppsRowNumArr[u], recordUpdatesColPos, 1, 1).setValue('Y');
      }
    }
  }
}
