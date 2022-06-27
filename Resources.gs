//add menu
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Birthday Reminder')
    .addItem('Send birthday reminder', 'sendEmail')
    .addToUi();
}

//connect to spreadsheet
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
