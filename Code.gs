// Sources: https://productivityspot.com/automatically-send-emails-from-google-sheets/
// https://yagisanatode.com/2019/03/07/google-apps-script-how-to-connect-a-button-to-a-function-in-google-sheets/
// Author: Chris Holt; December 14, 2020;

function safetyTest(ui, count) {
  var result = ui.prompt(
      'Are you sure you want to send ' + count + ' emails?',
      'Type the exact number shown to continue:',
      ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();

  if (button == ui.Button.OK) {
    // User clicked "OK".
    if(text == count){
      // User entered the right number
      ui.alert('Now sending ' + count + ' emails...');
      return true;
    } else {
      // User entered the wrong number
      ui.alert('No emails will be sent!');
      return false;
    }
    return false; // This line should never execute      
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('No emails will be sent!');
    return false;
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('No emails will be sent!');
    return false;
  } else {
    // Not sure what happened here, don't email anyone!
    ui.alert('No emails will be sent!');
    return false;
  }
  return false; // err on the side of NOT sending emails
}


function emailBlast(isDryRun, emailAddress, subject, message){
  if(isDryRun){
    // Dry Run so don't send an email.
    return false;
  } else if (!isDryRun){
    // NOT a Dry Run, send the email.
    MailApp.sendEmail(emailAddress, subject, message);
    return true;
  }
}

function sendEmails(isDryRun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var recipientSheet=ss.getSheetByName('Recipients');
  var templateSheet=ss.getSheetByName('Template');
  var subject = templateSheet.getRange(2,1).getValue();
  var count=recipientSheet.getLastRow();

  // Double check that the user meant to send this many emails
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var safetyCheck = safetyTest(ui, count - 1);

  if(isDryRun){
    // Post a summary message to the screen if this is a Dry Run.
    if(safetyCheck)
      var scmessage = 'Passed'
    else
      var scmessage = 'Failed'
    ui.alert('Dry Run Summary\n--------\n' + '\nEmails Found: ' + (count -1) + '\nSafetyTest Result: ' + scmessage);
  }

  if(safetyCheck){
    for (var i = 2; i < count+1 ; i++ ) {
      var emailAddress = recipientSheet.getRange(i,2).getValue();
      var name=recipientSheet.getRange(i,1).getValue();
      var trackingID=recipientSheet.getRange(i,3).getValue();
      var trackingLink=recipientSheet.getRange(i,4).getValue();
      var message = templateSheet.getRange(2,2).getValue();

      message=message.replace("<name>",name);
      message=message.replace("<trackingID>",trackingID);
      message=message.replace("<trackingLink>",trackingLink);

      emailBlast(isDryRun, emailAddress, subject, message);
    
      if((i % (Math.ceil(Math.random()*count+1) % 5) == (i % 5))){
        // Get a random number, MOD it by 5, if that matches the current index MOD 5, then update the progress bar, otherwise be silent
        // This makes sure that we see periodic updates to the UI, but doesn't force a sleep() between every execution
        updateProgressBar(i-1,count-1); // -1 is to adjust for the first line which is headers
        SpreadsheetApp.flush();
        Utilities.sleep(1000);
        }
      }
    updateProgressBar(i-2,count-1); // i-2 because it gets+1 to fall out of the loop
  }
}

function updateProgressBar(index,total) {
  template = HtmlService.createTemplateFromFile('StatusPage');
  template.index = index;
  template.total = total;
  var html = template.evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function dryRun(){
  sendEmails(true);
}

function wetRun(){
  sendEmails(false);
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Send Emails')
      .addItem('Initialize', 'onOpen')
      .addItem('Dry Run', 'dryRun')
      .addItem('Start Processing', 'wetRun')
      .addToUi();
}

