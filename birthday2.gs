const sheet = SpreadsheetApp.getActiveSpreadsheet.getSheetByName("Birthdays");
const data = sheet.getDataRange().getValues();
const lastCol = sheet.getLastColumn();
