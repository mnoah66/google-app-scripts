function sendEmail(e) {
  //setup function
  var ActiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  var StartRow = 1;
  var RowRange = ActiveSheet.getLastRow() - StartRow + 1;
  var MainData = ActiveSheet.getRange(StartRow,1,RowRange,5);
  var AllValues = MainData.getValues();
  
  
  var EmailActiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings'); 
  var EmailSheetStartRow = 1;
  var EmailRowRange = EmailActiveSheet.getLastRow() - StartRow + 1;
  var EmailData = EmailActiveSheet.getRange(EmailSheetStartRow, 1, EmailRowRange, 2);
  var AllValuesEmail = EmailData.getValues(); 
  
  //Build the email object
  var emailObj = new Object()
  for (i in AllValuesEmail) {
    var thisRow = AllValuesEmail[i];
    var program = thisRow[0];
    var email = thisRow[1];
    emailObj[program] = email;
  };
  
  //iterate loop
  for (i in AllValues) {
    //set current row
    var CurrentRow = AllValues[i];
    var scheduledDate = new Date(CurrentRow[3]);
    //Logger.log("outside of If statement: ",scheduledDate);
    // If todays date is greater than the date in column 4, I don't even want to evaluate it (No reason to send again)
    
    if (compareDates(scheduledDate) && (CurrentRow[4] != "sent") ) {
      //set subject line
      var Subject = "Training Reminder - " + CurrentRow[1];
      
      // format the scheduled date
      var scheduledDate = Utilities.formatDate(scheduledDate, "GMT", "MM/dd/yyyy")
      //set HTML template for information
      var message = 
            "<p>Staff <strong>" + CurrentRow[1] + "</strong> is scheduled for <strong>" + CurrentRow[2] + "</strong> on <strong>" + scheduledDate + "</strong><br><br>"  
              
      var signature = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings').getRange('F2').getValue();
     
      //mark row as "sent"
      var setRow = parseInt(i) + StartRow;
      ActiveSheet.getRange(setRow, 5).setValue("sent");
      
      //send the actual email  
      MailApp.sendEmail({
        to: emailObj[CurrentRow[0]],
        cc: "",
        name: "Training Reminder",
        subject: Subject,
        htmlBody: message + signature,
      });
    } // End of if statement
  } // End of iteration loop
  
  
} // End of the sendEmail() function

function compareDates(d) {
  var todayDate = new Date();
  if (todayDate > d) {
    return false;
  } else {
  var days = parseInt(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings').getRange('F3').getValue());
  if ((todayDate.setDate(todayDate.getDate()+days)) > d) {
    return true;
  }
    }
  
  
}