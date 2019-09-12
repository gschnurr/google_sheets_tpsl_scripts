/*
1) add safety in place to make sure that the sheetID in the response function is updated
2) create a function to run in the response function that will grab the answers and place them where they need to be in the spreadsheet

**/



function pp_form_gen() {

  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var vendorCatArr = ['Agencies', 'Commercial Partners', 'Credit Reference and Fraud Agencies',
  'Customer Service Outsourcing', 'Financial Products', 'General', 'Legal', 'Marketing and PR',
  'Operational Services', 'Payment Processors'];

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

  var ppe = ss.getSheetByName('PayPal Extract');
  var ppeLr = ppe.getLastRow();
  var ppeLc = ppe.getLastColumn();
  var ppeTitleColumnArr = ppe.getRange(1, 1, 1, ppeLc).getValues();

  var ppeOned = flatten_arr(ppeTitleColumnArr);

  var busSysOwnColPos = find_col(ppeOned, 'Business System Owner');
  var busSysOwnerArr = ppe.getRange(2, busSysOwnColPos, ppeLr, 1).getValues();
  var bsoOned = flatten_arr(busSysOwnerArr);
  bsoOned.pop();

  var appNameColPos = find_col(ppeOned, 'Application');
  var appNameArr = ppe.getRange(2, appNameColPos, ppeLr, 1).getValues();
  var anOned = flatten_arr(appNameArr);

  //this is an array of the row number of the applications owned by the currentBoInLoop
  //creates an arr of all col pos of columns to be updated
  var gdprBoColPosArr = [];
  for (var c = 0; c < gdprBoTColArr.length; c++) {
    var gdprBoColPos = find_col(ppeOned, gdprBoTColArr[c]);
    gdprBoColPosArr.push(gdprBoColPos);
  }
  //an array of the bo's who have been looped through to prevent creating multiple forms for one user
  var boFormComArr = ['placeholder', ''];
  //loop through the business owners arr
  for (var b = 0; b < bsoOned.length; b++) {
    Logger.log('BOArry Loop, BO = ' + bsoOned[b] + ' b = ' + b);
    Logger.log(boFormComArr.length + ' number of forms have been sent.')
    //check the current iteration of the bo arr against the bo complete? arr
    for (var c = 0; c < boFormComArr.length; c++) {
      Logger.log('BOs that have been used are ' + boFormComArr);
      // if the first current item in the completed array = the current bo then go to the next bo
      if (boFormComArr[c] == bsoOned[b]) {
        Logger.log('BOCompArr Loop - BO Has already been used, BO = ' + bsoOned[b] + ' c = ' + c + ' length = ' + boFormComArr.length);
        break;
      }
      else if (boFormComArr[c] != bsoOned[b] && c != (boFormComArr.length - 1)){
        Logger.log('BOCompArr Loop - BO is not in ARR but we are not at the end of the array yet, BO = ' + bsoOned[b] + ' c = ' + c + ' length = ' + boFormComArr.length);
        continue;
      }
      else if (boFormComArr[c] != bsoOned[b] && c == (boFormComArr.length - 1)) {
        Logger.log('BOCompArr Loop - BO Not Found, create form. BO = ' + bsoOned[b] + ' c = ' + c + ' length = ' + boFormComArr.length);
        boFormComArr.unshift(bsoOned[b]);
        var currentBoInLoop = bsoOned[b];
        var userUpdatesForm = FormApp.create(currentBoInLoop + ' Applications');
        //this is an array of the row number of the applications owned by the currentBoInLoop
        var boAppsRowNumArr = [];
        var boAppsArr = [];
        for (var f = 0; f < bsoOned.length; f++) {
          if (bsoOned[f] == currentBoInLoop) {
            var cwuAppRow = f + 2;
            boAppsRowNumArr.push(cwuAppRow);
            var appNameCell = ppe.getRange(cwuAppRow, appNameColPos, 1, 1).getValue();
            boAppsArr.push(appNameCell);
          }
          else {
            continue;
          }
        }
        Logger.log('Form For ' + bsoOned[b] + ' created ' + bsoOned[b] + ' has ' + boAppsArr.length + ' Applications');
        for (var u = 0; u < boAppsRowNumArr.length; u++) {
          var appToUpdate = boAppsArr[u];
          userUpdatesForm.addPageBreakItem().setTitle(appToUpdate);
          Logger.log('the first application is ' + appToUpdate);
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
          }
          Logger.log('All items are added for ' + appToUpdate);
        }
        Logger.log('All applications are added to the form. ' + u + ' applications were added.'');
        ScriptApp.newTrigger('on_Form_Sub_Bo_Trigger')
          .forForm(userUpdatesForm)
          .onFormSubmit()
          .create();
        var responseUrl = userUpdatesForm.getPublishedUrl();
          var emailTo = currentBoInLoop;
          var subject = 'test';
          var options = {}
          options.htmlBody = "Hi Everyone-" +'<br />'+'<br />'+ "Here\'s the " + '<a href=\"' + responseUrl + '">form URL</a>';
          MailApp.sendEmail(emailTo, subject, '', options);
        Logger.log('Email sent');
        break;
      }
      else {
        ui.alert('OOPs: something went wrong. Please contact and administrator.');
      }
    }
  }
}


function on_Form_Sub_Bo_Trigger(e) {
  // get the trigger id passed from the event and use that to get the form id
  var triggerId = e.triggerUid;
  var formId = get_file_by_trigger_id(triggerId);
  var form = FormApp.openById(formId);
  var test = get_form_responses_formatted_Arr(form);
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
  var ppe = ss.getSheetByName('PayPal Extract');
  MailApp.sendEmail('gibson.schnurr@izettle.com', 'did this work', test);
}

function get_file_by_trigger_id(triggerId) {
  var triggers = ScriptApp.getProjectTriggers();
  for(var i = 0; i < triggers.length; i++){
    if (triggers[i].getUniqueId() == triggerId) {
      return triggers[i].getTriggerSourceId();
    }
  }
}

// ideally this function will return an array of [[formitem, formresponse]] not sure actually might need to find a way to match things up and get a concise return
function get_form_responses_formatted_Arr(form) {
  var itemArr = form.getItems();
  return itemArr[0].getId();
}

function update_app_info_in_ss() {

}
