function sendEmailFullAuto(){
    //connect to spreadsheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LGS Birthdays v2");
    const data = sheet.getDataRange().getValues();
    const lastCol = sheet.getLastColumn();
    const headersText = sheet.getRange(1, 1, 1, lastCol).getValues();
  
    //get header row
    const headerObj = {};
    for (let i = 0; i < headersText[0].length; i++) {
      const header = headersText[0][i];
      headerObj[`${header}_Col`] = data[0].indexOf(header);
    }
    console.log("headers:", headerObj)
  
    //create email list
    const emailList = generateEmailList_(data, headerObj);
    console.log("full email list:", emailList)
  
    for(i=1; i<data.length; i++){
      //check for birthday dates
      let birthday = data[i][headerObj['birthday_Col']];
      console.log("birthday:", birthday)
      if(birthday == ""){
        continue
      }
  
      //check if today is one week before birthday
      let today = new Date()
      let todayFormat = new Date(today.getFullYear(), today.getMonth(), today.getDate())
      const birthdayMinus7 = new Date(today.getFullYear(), birthday.getMonth(), birthday.getDate() - 7)
      console.log("today:", todayFormat)
      console.log("birthday-7:", birthdayMinus7)

      const link = data[i][headerObj['link_Col']];
  
      //if today is one week before birthday and link is not blank...
      if(todayFormat.getTime() == birthdayMinus7.getTime() && link){
        console.warn("Send yosetti link!", link)
        const name = data[i][headerObj['name_Col']];
        const email = data[i][headerObj['email_Col']];
        const gender = data[i][headerObj['gender_Col']];
        const status = data[i][headerObj['status_Col']];
  
        //create new email list
        const birthdayIndex = emailList.indexOf(email);
        const newEmailList = generateNewEmailList_(emailList, birthdayIndex)
        console.log("New email list:", newEmailList)
        
        const dueDate = getDueDate_(birthday);
        console.log("Due:", dueDate)
        
        //email template
        const dueDateFormat = Utilities.formatDate(dueDate, "Asia/Tokyo", 'MMMM dd')
        const birthdayFormat = Utilities.formatDate(birthday, "Asia/Tokyo", 'MMMM dd')
        const emailTemplate = "Birthday mail";
        const templateIndex = HtmlService.createTemplateFromFile(emailTemplate);
        templateIndex.name = name;
        templateIndex.gender = gender;
        templateIndex.birthdayFormat = birthdayFormat;
        templateIndex.dueDate = dueDateFormat;
        templateIndex.link = link;

        const emailBody = templateIndex.evaluate().getContent();

        //emailer
        const myEmail = Session.getActiveUser().getEmail()
        const subject = "RRO: Request for Birthday Messages to " + name + " by " + dueDateFormat;
        const body = "";
        const options = { 
            bcc: newEmailList,
            htmlBody: emailBody,
            }
        try{
          GmailApp.sendEmail(myEmail, subject, body, options);
          console.log("Email sent!")
          //sheet.getRange(i+1, status+1).setValue("Email sent");
        } catch (error){
          console.error("Email error:", error);
          //sheet.getRange(i+1, status+1).setValue(error);
        }
      }
    }
  }

function generateEmailList_(data, headerObj) {
    let emailList = [];
    for(i=1; i<data.length; i++){
      emailList.push(data[i][headerObj['email_Col']]);
    }
    return emailList;
}

function generateNewEmailList_(emailList, index){
  emailList.splice(index, 1)
  const newEmailList = emailList.join()
  return newEmailList
}

function getDueDate_(birthday) {
  const today = new Date()
  let dueDate = new Date(today.getFullYear(), birthday.getMonth(), birthday.getDate()-2)
  if (dueDate.getDay() === 6) { // 6 = Sat
    dueDate = new Date(today.getFullYear(), dueDate.getMonth(), dueDate.getDate()-1)
  } else if (dueDate.getDay() === 0) { // 0 = Sun
    dueDate = new Date(today.getFullYear(), dueDate.getMonth(), dueDate.getDate()-2)
  }
  return dueDate;
}