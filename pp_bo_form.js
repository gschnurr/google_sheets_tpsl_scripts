function pp_form_gen() {

//need to seperate variables into static and dynamic
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var ppe = ss.getSheetByName('PayPal Extract');
  var ppeLr = ppe.getLastRow(); //dynamic
  var ppeLc = ppe.getLastColumn(); //dynamic
  var ppeTitleColumnArr = ppe.getRange(1, 1, 1, ppeLc).getValues();

  //this array contains the choices for vendor category
  var vendorCatArr = ['Agencies', 'Commercial Partners', 'Credit Reference and Fraud Agencies',
  'Customer Service Outsourcing', 'Financial Products', 'General', 'Legal', 'Marketing and PR',
  'Operational Services', 'Payment Processors'];

  //this array contains the choices for purpose
  var purposeCatArr = ['To provide our services and products, to fulfill relevant agreements with our merchants and to otherwise administer our business relationship with our merchants.',
  'To confirm your identity and verify our merchant’s personal and contact details.',
  'To prove that transactions have been executed.', 'To establish, exercise or defend a legal claim or collection procedures.',
  'To comply with internal procedures.',
  'To administer payment for products and/or services and the customer relationship i.e. to carry out our obligations arising from any contracts entered into between us and the merchant and to provide you with the information, products, and services that you request from us.',
  'To assess which payment options and payment services to offer to our merchants, for example by carrying out internal and external credit assessments.',
  'For customer analysis, to administer iZettle´s services, and for internal operations, including troubleshooting, data analysis, testing, research, and statistical purposes.',
  'To ensure that content is presented in the most effective way for our merchants and their device.',
  'To prevent misuse of iZettle´s services as part of our efforts to keep our services safe and secure.', 'To carry out risk analysis, fraud prevention, and risk management.',
  'To improve our services and for general business development purposes, such as improving credit risk models in order to e.g. minimize fraud, develop new products and features and explore new business opportunities.',
  'Marketing, product and customer analysis. This processing forms the basis for marketing, process and system development, including testing. This is to improve our product range and to optimize our customer offering.',
  'To comply with applicable laws, such as anti-money laundering and bookkeeping laws and regulatory capital adequacy requirements and rules issued by our designated banks and relevant card networks. This means that we process personal data for know-your-customer (“KYC”) ' +
  'requirements, to prevent, detect and investigate money laundering, terrorist financing, and fraud. We also carry out sanction screening, report to tax authorities, police enforcement authorities, enforcement authorities, supervisory authorities.',
  'To be able to administer participation in competitions and/or events.',
  'Risk management obligations such as credit performance and quality, insurance risks and compliance with capital adequacy requirements under applicable law.',
  'Risk management obligations such as credit performance and quality, insurance risks and compliance with capital adequacy requirements under applicable law.',
  'To administer payments carried out by using our services from a Merchant.', 'To communicate with our merchants in relation to our services.'];

  //this loops checks all of the current sheets in the spreadsheet for the spreadsheet savePointSheet
  var savePointSheet = 'PayPal Extract Save';
  logs_tst('All static variables have been initialized.');

  for (var s = 0; s < sheets.length; s++){
    curSheetName = sheets[s].getSheetName();
    if (curSheetName == savePointSheet) {
      logs_tst('Sheet with sheet name ' + savePointSheet + ' already exists.');
      break;
    }
    else if (curSheetName != savePointSheet && s != (sheets.length - 1)) {
      continue;
    }
    else if (curSheetName != savePointSheet && s == (sheets.length - 1)){
      spreadsheet.insertSheet(1);
      spreadsheet.getActiveSheet().setName('PayPal Extract Save');
      var ppeSave = ss.getSheetByName('PayPal Extract Save');
      ppeSave.getRange('A1').activate();
      ppe.getRange(1, 1, 1, ppeLc).copyTo(ppeSave.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      ppeSave.getRange('A1').activate();
      ppe.getRange(1, 1, 1, ppeLc).copyTo(ppeSave.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      ppeSave.setFrozenRows(1);
      logs_tst('Sheet with sheet name ' + savePointSheet + ' created and formatted.');
    }
    else {
      ui.alert('OOPs: something went wrong. Please contact and administrator.');
    }
  }

  var ppeSave = ss.getSheetByName('PayPal Extract Save');
  var ppeSaveLr = ppeSave.getLastRow();
  var ppeSaveLc = ppeSave.getLastColumn();
  var ppeSaveLrPOne = ppeSaveLr + 1;
  logs_tst('ppeSave variables initialized.');

  //creates 1d array of the title row values for the ppe spreadsheet
  var ppeOned = flatten_arr(ppeTitleColumnArr);
  logs_tst('PPE title row array created.');
  //finds the position of the business system owner column grabs all of the data in that column and creates a 1d array and removes the last blank value
  var busSysOwnColPos = find_col(ppeOned, 'Business System Owner');
  var busSysOwnerArr = ppe.getRange(2, busSysOwnColPos, ppeLr, 1).getValues();
  var bsoOned = flatten_arr(busSysOwnerArr);
  bsoOned.pop();
  logs_tst('One dimensional BO array created.');

  //creates an arr of all col pos of columns to be placed in the form and updated
  var gdprBoColPosArr = [];
  for (var c = 0; c < gdprBoTColArr.length; c++) {
    var gdprBoColPos = find_col(ppeOned, gdprBoTColArr[c]);
    gdprBoColPosArr.push(gdprBoColPos);
  }
  logs_tst('GDPR Column positions to be updated from initial static array have been found.');

  //an array of the bo's who have been looped through to prevent creating multiple forms for one user
  var boFormComArr = [''];
  var bson = 0;
  var activeTriggerPrompt = ui.prompt('How many active form response triggers are there for this script?').getResponseText();
  var userInFormsSentNum = Number(activeTriggerPrompt); // this will be the response from a prompt about the numebr of active triggers left if there is a save point
  var numFormsSent = (userInFormsSentNum + boFormComArr.length - 1); //needs to be inside loop the as well this is the first initilization
//beginning of actual function
  while (numFormsSent < 20) {
    logs_tst('Number of forms sent = ' + numFormsSent + '.');
    //create the one dimensional array of all the business owners the BO column for this iteration
    var ppeLr = ppe.getLastRow();
    var busSysOwnerArr = ppe.getRange(2, busSysOwnColPos, ppeLr, 1).getValues();
    var bsoOned = flatten_arr(busSysOwnerArr);
    bsoOned.pop();
    //create a one dimensional array of all the applications still in the ppe for this iteration
    var appNameColPos = find_col(ppeOned, 'Application');
    var appNameArr = ppe.getRange(2, appNameColPos, ppeLr, 1).getValues();
    var appNameOned = flatten_arr(appNameArr);

    if (bsoOned.length > 0 && bson < bsoOned.length) {
      var boToCheck = bsoOned[bson];
      logs_tst('BOArry Loop prior to check, BO = ' + boToCheck);
    }
    else {
      logs_tst('bsoOned.length = ' + bsoOned.length + '. and/or bson = ' + bson + '.');
      var compLogs = Logger.getLog();
      MailApp.sendEmail('gibson.schnurr@izettle.com', 'PP BO Form Script', compLogs);
      ui.alert('OPERATION COMPLETE: There are no unique business system owners left.')
      return;
    }
    for (var c = 0; c < boFormComArr.length; c++) {
      logs_tst('BOs that have been used are ' + boFormComArr);
      // if the first current item in the completed array = the current bo then go to the next bo
      if (boFormComArr[c] == boToCheck) {
        logs_tst('BOCompArr Loop - BO Has already been used, BO = ' + boToCheck + ' c = ' + c + ' length = ' + boFormComArr.length);
        bson++;
        logs_tst('bson incrimented to ' + bson);
        break;
      }
      else if (boFormComArr[c] != boToCheck && c != (boFormComArr.length - 1)){
        logs_tst('BOCompArr Loop - BO is not in ARR but we are not at the end of the array yet, BO = ' + boToCheck + ' c = ' + c + ' length = ' + boFormComArr.length);
        continue;
      }
      else if (boFormComArr[c] != boToCheck && c == (boFormComArr.length - 1)) {
        logs_tst('BOCompArr Loop - BO Not Found, create form. BO = ' + boToCheck + ' c = ' + c + ' length = ' + boFormComArr.length);
        bson = 0;
        logs_tst('bson reset to zero.')
        //defining the currentBO Email as a variable and adding that BO to the already used array
        var boOfCurFormCre = boToCheck;
        boFormComArr.unshift(boOfCurFormCre);
        var userUpdatesForm = FormApp.create(boOfCurFormCre + ' Applications');
        userUpdatesForm.setTitle(boOfCurFormCre + ' Applications');
        //initial empty arrays for the currentBos app information
        var boAppsRowNumArr = [];
        var boAppsArr = [];
        //push information into those arrays to be used to locate the current BOs applications via the below loop
        for (var f = 0; f < bsoOned.length; f++) {
          if (bsoOned[f] == boOfCurFormCre) {
            var cwuAppRow = f + 2;
            boAppsRowNumArr.push(cwuAppRow);
            var appNameCell = ppe.getRange(cwuAppRow, appNameColPos, 1, 1).getValue();
            boAppsArr.push(appNameCell);
          }
          else {
            continue;
          }
        }
        logs_tst('Form For ' + boOfCurFormCre + ' created ' + boOfCurFormCre + ' has ' + boAppsArr.length + ' Applications');
        for (var u = 0; u < boAppsRowNumArr.length; u++) {
          var appToUpdate = boAppsArr[u];
          userUpdatesForm.addPageBreakItem().setTitle(appToUpdate);
          logs_tst('the first application is ' + appToUpdate);
          for (var v = 0; v < gdprBoColPosArr.length; v++) {
            var colTitle = gdprBoTColArr[v];
            var colHelpText = ppe.getRange(1, gdprBoColPosArr[v], 1, 1).getNotes();
            var currentInfo = ppe.getRange(boAppsRowNumArr[u], gdprBoColPosArr[v], 1, 1).getValue();
            var currentInfoCheck;
            if (currentInfo == '') {
            currentInfoCheck = '<BLANK>';
            }
            else {
              currentInfoCheck = currentInfo;
            }
            //create an if statement here depending on the column if it is yes no make it a yes know response
            if (colTitle == 'GDPR Data (Y,N)') {
              var multiChoice = userUpdatesForm.addMultipleChoiceItem().setRequired(false);
                multiChoice.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
                  .setChoices([
                    multiChoice.createChoice('Y'),
                    multiChoice.createChoice('N'),
                  ]);
            }
            else if (colTitle == 'Employee Data') {
              var multiChoice = userUpdatesForm.addMultipleChoiceItem().setRequired(false);
                multiChoice.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
                  .setChoices([
                    multiChoice.createChoice('Yes'),
                    multiChoice.createChoice('No'),
                  ]);
            }
            else if (colTitle == 'End Customer Data') {
              var multiChoice = userUpdatesForm.addMultipleChoiceItem().setRequired(false);
                multiChoice.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
                  .setChoices([
                    multiChoice.createChoice('Yes'),
                    multiChoice.createChoice('No'),
                  ]);
            }
            else if (colTitle == 'Merchant Data') {
              var multiChoice = userUpdatesForm.addMultipleChoiceItem().setRequired(false);
                multiChoice.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
                  .setChoices([
                    multiChoice.createChoice('Yes'),
                    multiChoice.createChoice('No'),
                  ]);
            }
            else if (colTitle == 'Vendor Category') {
              var multiChoice = userUpdatesForm.addMultipleChoiceItem().setRequired(false);
              var mcOptions = [];
              for (var mc = 0; mc < vendorCatArr.length; mc++) {
                var vcChoice = multiChoice.createChoice(vendorCatArr[mc]);
                mcOptions.push(vcChoice);
              }
                multiChoice.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
                  .setChoices(mcOptions);
            }
            else if (colTitle == 'Purpose') {
              var multiChoice = userUpdatesForm.addMultipleChoiceItem().setRequired(false);
              var mcOptions = [];
              for (var mc = 0; mc < purposeCatArr.length; mc++) {
                var purpChoice = multiChoice.createChoice(purposeCatArr[mc]);
                mcOptions.push(purpChoice);
              }
                multiChoice.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
                  .setChoices(mcOptions);
            }
            else if (colTitle == 'Data Disclosed') {
              var textItem = userUpdatesForm.addTextItem().setRequired(false);
                textItem.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
            }
            else if (colTitle == 'Data shared with third party? (Y,N,N/A)') {
              var multiChoice = userUpdatesForm.addMultipleChoiceItem().setRequired(false);
                multiChoice.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
                  .setChoices([
                    multiChoice.createChoice('Yes'),
                    multiChoice.createChoice('No'),
                    multiChoice.createChoice('N/A'),
                  ]);
            }
            else if (colTitle == 'Headquarter location') {
              var textItem = userUpdatesForm.addTextItem().setRequired(false);
                textItem.setTitle(colTitle)
                  .setHelpText(colHelpText + ' The current information is ' + currentInfoCheck + '.')
            }
          } //end of the loop for creating gdpr form elements
          logs_tst('All items are added for ' + appToUpdate);
        } //end of the loop for currentBO applications loop
        logs_tst('All applications are added to the form. ' + u + ' applications were added.');
        ScriptApp.newTrigger('on_Form_Sub_Bo_Trigger')
          .forForm(userUpdatesForm)
          .onFormSubmit()
          .create();
        var responseUrl = userUpdatesForm.getPublishedUrl();
          var emailTo = boOfCurFormCre;
          var subject = 'Test - Please ignore this, this is just for testing purposes.';
          var options = {}
          options.htmlBody = "Hi Everyone-" +'<br />'+'<br />'+ "Here\'s the " + '<a href=\"' + responseUrl + '">form URL</a>';
          MailApp.sendEmail(emailTo, subject, '', options);
        logs_tst('Email sent, copy and save initiated.');

        var appDelRowAdj = 0;
        for (var z = 0; z < boAppsArr.length; z++ ) {
          for (var m = 0; m < appNameArr.length; m ++ ) {
            if (appNameArr[m] == boAppsArr[z]) {
              var ppeSaveLr = ppeSave.getLastRow();
              var ppeSaveLrPOne = ppeSaveLr + 1;
              var appToRemRow = (m + 2 - appDelRowAdj);
              var appRangeToCopy = ppe.getRange(appToRemRow, 1, 1, ppeLc).getValues();
              ppeSave.getRange(ppeSaveLrPOne, 1, 1, ppeSaveLc).setValues(appRangeToCopy);
              ppe.deleteRow(appToRemRow);
              appDelRowAdj++;
              logs_tst('appDelRowAdj incremented to ' + appDelRowAdj);
              logs_tst(boAppsArr[z] + ' deleted and saved.');
            }
            else {
              continue;
            }
          }
        }
        var numFormsSent = (userInFormsSentNum + boFormComArr.length - 1); //we want to reevaluate the length of the check arry each iteration
        logs_tst('Number of forms sent = ' + numFormsSent + '.');
        logs_tst('All applications for ' + boOfCurFormCre + ' have been deleted.');
        break;
      } //end of the elseif for creating a new form for user not found in the check array
      else {
        ui.alert('OOPs: something went wrong. Please contact and administrator.');
      } // end of the else right above this
    } //end of the forloop for the check against boFormComArr
  } //end of the while loops
  logs_tst('The max number of triggers have been reached.');
  var trigLimReachedLogs = Logger.getLog();
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'PP BO Form Script', trigLimReachedLogs);
  ui.alert('OPERATION COMPLETE/WARNING: Maximum number of triggers exceeded.');
} //end of the function

function on_Form_Sub_Bo_Trigger(e) {
  // get the trigger id passed from the event and use that to get the form id
  var triggerId = e.triggerUid;
  var formId = get_file_by_trigger_id(triggerId);
  var formPpResp = FormApp.openById(formId);
  logs_tst('A form has been submitted, the Form ID is ' + formId + ' and the trigger id is ' + triggerId);
  // the below might not work depending on how drive orders its files
  // begin searching through google drive files for files containing the below text in the title
  var files = DriveApp.searchFiles('title contains "PayPal Extract"');
  //push those file's ids into an array
  var ppExtracts = [];
  while (files.hasNext()) {
    var file = files.next();
    ppExtracts.push(file.getId());
  }
  //grab the latest file id containing that text and open it in the background
  var latestPpSsValue = (ppExtracts.length - 1);
  var ssId = ppExtracts[latestPpSsValue];
  var ss = SpreadsheetApp.openById(ssId);
  var ppeSave = ss.getSheetByName('PayPal Extract Save');
  logs_tst('The spreadsheet to be opened is ' + ss);

  getResp_update(formPpResp, ppeSave);
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'did this work', test);
}

function get_file_by_trigger_id(triggerId) {
  var triggers = ScriptApp.getProjectTriggers();
  for(var t = 0; t < triggers.length; t++){
    if (triggers[t].getUniqueId() == triggerId) {
      return triggers[t].getTriggerSourceId();
    }
  }
}

function getResp_update(form, updateSheet) {
  var itemArr = form.getItems();
  logs_tst('The Form is ' + form.getTitle());
  logs_tst('The update sheet is ' + updateSheet.getSheetName());
  logs_tst('The itemArr is ' + itemArr);
  var formResponses = form.getResponses();
  var ppeSave = updateSheet;
  var ppeSaveLc = ppeSave.getLastColumn();
  var ppeSaveLr = ppeSave.getLastRow();
  var ppeSaveTitleColumnArr = ppeSave.getRange(1, 1, 1, ppeSaveLc).getValues();
  var ppeSaveTitleColumnArrOned = flatten_arr(ppeSaveTitleColumnArr);
  var colTitleToFind = 'Application';
  var appNameTitleColPos = find_col(ppeSaveTitleColumnArrOned, colTitleToFind);
  var ppeSaveAppArr = ppeSave.getRange(2, appNameTitleColPos, ppeSaveLr, 1).getValues();
  var ppeSaveAppArrOned = flatten_arr(ppeSaveAppArr);

  for (var y = 0; y < itemArr.length; y++) {
    logs_tst('Y is ' + y + ' at the start of this for loop iteration.');
    var curItemType = itemArr[y].getType();
    logs_tst('The item type of the item in the item array loop is ' + curItemType);
    if (curItemType == FormApp.ItemType.PAGE_BREAK) {
      var curAppName = itemArr[y].getTitle();
      logs_tst('The current item type equals page break. The applicaiton is ' + curAppName);
      var curItemAppRow = find_row(ppeSaveAppArrOned, curAppName);
      logs_tst(curAppName + ' is in row ' + curItemAppRow + ' in the ppeSave sheet.');
      y++;
      logs_tst('Y is ' + y + ' before the while loop for this iteration.');
      var firAppItemType = itemArr[y].getType();
      logs_tst('The next item type in the array is ' + firAppItemType);
      if (firAppItemType == FormApp.ItemType.MULTIPLE_CHOICE || firAppItemType == FormApp.ItemType.TEXT ) {
        var nextItemType = firAppItemType;
        while (nextItemType == FormApp.ItemType.MULTIPLE_CHOICE || nextItemType == FormApp.ItemType.TEXT) {
          logs_tst('Beginning of the while loop for this iteration. Y = ' + y + ' nextItemType = ' + nextItemType);
          var respItemColTi = itemArr[y].getTitle(); //column title
          var respItemColTiPos = find_col(ppeSaveTitleColumnArrOned, respItemColTi);
          var respItemId = itemArr[y].getId();
          logs_tst('The response item id = ' + respItemId + ' , before the for loop of looping through the form responses');
          var lastFormSub = formResponses.length - 1;
          var formResponseNewest = formResponses[lastFormSub];
          var itemResponseInst = formResponseNewest.getItemResponses();
          for (var ir = 0; ir < itemResponseInst.length; ir++) {
            var curItemResp = itemResponseInst[ir];
            var itemRespIdforCurRespI = curItemResp.getItem().getId();
            logs_tst('ID of respitem id does it match? ' + itemRespIdforCurRespI);
            if (itemRespIdforCurRespI == respItemId) {
              var respItemResp = curItemResp.getResponse();
              logs_tst('The response for this item is ' + respItemResp);
              break;
            }
            else if (itemRespIdforCurRespI != respItemId && ir != (itemResponseInst.length - 1)) {
              continue;
            }
            else if (itemRespIdforCurRespI != respItemId && ir == (itemResponseInst.length - 1)) {
              var respItemResp = ppeSave.getRange(curItemAppRow, respItemColTiPos, 1, 1).getValue();
              logs_tst('The default value has been entered in the cell');
            }
            else {
              MailApp.sendEmail('gibson.schnurr@izettle.com', 'something went wrong', 'Something went wrong in the for loop for checking IDs');
              return;
            }
          }
          logs_tst('The coordinates for the cell to be replaced are (in x,y format) ' + '(' + respItemColTiPos + ',' + curItemAppRow + ')' +
          ' The value that will be set in that cell is ' + respItemResp);
          ppeSave.getRange(curItemAppRow, respItemColTiPos, 1, 1).setValue(respItemResp);
          y++;
          if (y < itemArr.length) {
            logs_tst('Y has been incremented in the while loop to ' + y);
            var nextItem = itemArr[y];
            var nextItemType = nextItem.getType();
            logs_tst('The next item type in the array is ' + nextItemType);
          }
          else {
            logs_tst('Y has been decremented. Y = ' + y);
            break;
          }
        }
        y--;
        logs_tst('The while loop has ended. the nextItemType variable = ' + nextItemType + ' Y = ' + y);
      }
      else {
        logs_tst('If statement before while loop did not pass, firAppItemType = ' + firAppItemType);
        continue;
      }
    }
    else {
      logs_tst('The the item failed the first if statement in the for loop. The item type that failed was ' + curItemType);
      continue;
    }
  }
  var compLogs1 = Logger.getLog();
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'On Form Sub Logs Comp Log', compLogs1);
  logs_tst('The for loop has ended.');
}
