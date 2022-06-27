function sendEmail() {
  //create email list
  const birthdaySheetName = 'LGS Birthdays';
  const birthdaySheet = spreadsheet.getSheetByName(birthdaySheetName);

  //combine all email addresses into one array
  const row = 2;
  const column = 6;
  const numRows = birthdaySheet.getDataRange().getLastRow();
  const numColumns = birthdaySheet.getDataRange().getLastColumn();
  let emailData = birthdaySheet.getRange(row, column, numRows, numColumns).getValues();
  let emailListAll = [];
  for (i = 0; i < emailData.length; i++) {
    let data = emailData[i][0];
    if (data !== '') {
      emailListAll.push(data);
    };
  }

  //define birthday person's email address and remove from email list
  const birthdayEmail = birthdaySheet.getRange("L1").getValue();
  const birthdayIndex = emailListAll.indexOf(birthdayEmail);
  emailListAll.splice(birthdayIndex,1);
  const newEmailList = emailListAll.join(); //creates a string of email addresses

  console.log(newEmailList);

  //email variables
  const birthdayName = birthdaySheet.getRange("K1").getValue();
  const gender = birthdaySheet.getRange("N1").getValue();
  const birthdayDate = birthdaySheet.getRange("M1").getValue();
  const formatBirthdayDate = Utilities.formatDate(birthdayDate, 'Asia/Tokyo', 'MMMM d');
  const dueDate = new Date(birthdaySheet.getRange("O1").getValue());
  const formatDueDate = Utilities.formatDate(dueDate,'Asia/Tokyo', 'MMMM d');

  const yosettiLink = birthdaySheet.getRange("P1").getValue();

  //email template
  const emailTemplate = "Birthday mail";
  const templateIndex = HtmlService.createTemplateFromFile(emailTemplate);
  templateIndex.birthdayName = birthdayName;
  templateIndex.gender = gender;
  templateIndex.formatBirthdayDate = formatBirthdayDate;
  templateIndex.formatDueDate = formatDueDate;
  templateIndex.yosettiLink = yosettiLink;

  const emailBody = templateIndex.evaluate().getContent();

  //emailer
  const subject = "RRO: Request for Birthday Messages to " + birthdayName + " by " +formatDueDate;
  const options = { 
    bcc: newEmailList,
    htmlBody: emailBody,
    name: "Birthday Emailing Committee" 
    }
  const body = "";//"Dear All. Sorry for making you my test subjects. Please ignore this email. -Daryl";
  MailApp.sendEmail("hori.himawari@link-gs.co.jp", subject, body, options);
  
}
