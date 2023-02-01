function addDueDatesToCal() {
  const calendarID = "your@email.com";
  const datesSheetName = 'Add to Calendar';
  const datesSheet = spreadsheet.getSheetByName(datesSheetName); 

  const eventCal = CalendarApp.getCalendarById(calendarID);

  const eventData = datesSheet.getRange("F2:H35").getValues();
  //should get the last row for this instead

  for (i = 0; i < eventData.length; i++) {
    let events = eventData[i];
    let date = events[0];
    let title = events[2];
    console.log(date, title) 
    //eventCal.createAllDayEvent(title, date);
  }
}