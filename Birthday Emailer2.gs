function sendEmailFullAuto(){
    //connect to spreadsheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    const data = sheet.getDataRange().getValues();
    const lastCol = sheet.getLastColumn();
    const headersText = sheet.getRange(1, 1, 1, lastCol).getValues();
  
    //get header row
    const headerObj = {};
    for (let i = 0; i < headersText[0].length; i++) {
      const header = headersText[0][i];
      headerObj[`${header}_Col`] = data[0].indexOf(header);
    }
    // console.log(headerObj)
  
    //create email list
    let emailList = [];
    for(a=1; a<data.length; a++){
      emailList.push(data[a][headerObj['email_Col']]);
    }
    console.log(emailList)
  
    for(i=1; i<data.length; i++){
      //check for birthday dates
      let birthday = data[i][headerObj['birthday_Col']];
      console.log("birthday:", birthday)
      if(birthday == ""){
        continue
      }
  
      //check if today is one week before birthday
      let today = new Date()
      today = new Date(today.getFullYear(), today.getMonth(), today.getDate())
      const birthdayMinus7 = new Date(today.getFullYear(), birthday.getMonth(), birthday.getDate() - 7)
      console.log("today:", today)
      console.log("birthday-7:", birthdayMinus7)

      const link = data[i][headerObj['link_Col']];
  
      //if today is one week before birthday and link is not blank...
      if(today.getTime() == birthdayMinus7.getTime() && link !== ""){
        console.warn("Send yosetti link!")
        const name = data[i][headerObj['name_Col']];
        const email = data[i][headerObj['email_Col']];
        const gender = data[i][headerObj['gender_Col']];
        
        const status = data[i][headerObj['status_Col']];
  
        //create new email list
        const birthdayIndex = emailList.indexOf(email);
        emailList.splice(birthdayIndex, 1);
        const newEmailList = emailList.join();
        console.log(newEmailList)
  
        const birthdayFormat = Utilities.formatDate(birthday, "Asia/Tokyo", 'MMMM dd')
        // console.log(birthdayFormat)
        let dueDate = new Date(today.getFullYear(), birthday.getMonth(), birthday.getDate()-2)
        if(dueDate.getDay() == 6){ // 6 = Sat
          dueDate = new Date(today.getFullYear(), dueDate.getMonth(), dueDate.getDate()-1)
        } else if(dueDate.getDay() == 0){ // 0 = Sun
          dueDate = new Date(today.getFullYear(), dueDate.getMonth(), dueDate.getDate()-2)
        }
        console.log("Due:", dueDate)
        const dueDateFormat = Utilities.formatDate(dueDate, "Asia/Tokyo", 'MMMM dd')
  
        //email template
        const emailTemplate = "Birthday Reminder Email";
        const templateIndex = HtmlService.createTemplateFromFile(emailTemplate);
        templateIndex.name = name;
        templateIndex.gender = gender;
        templateIndex.birthdayFormat = birthdayFormat;
        templateIndex.dueDate = dueDateFormat;
        templateIndex.link = link;

        const emailBody = templateIndex.evaluate().getContent();

        //emailer
        const subject = "RRO: Request for Birthday Messages to " + name + " by " + dueDateFormat;
        const options = { 
            bcc: newEmailList,
            htmlBody: emailBody,
            }
        try{
          //GmailApp.sendEmail(<YOUR EMAIL>, subject, htmlBody, options);
          sheet.getRange(i+1, status+1).setValue("Email sent");
        } catch (error){
          console.error(error);
          sheet.getRange(i+1, status+1).setValue(error);
        }
            
      }
    }
  }
  