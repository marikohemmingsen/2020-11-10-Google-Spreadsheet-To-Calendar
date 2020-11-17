/* カレンダーIDを入力してください。*/
var CALENDAR_ID = '';

/* カレンダーにデータを送信する、タブの名前を入力してください。*/
var sheetName = 'シート1';

/* 各項目のカラムの番号を入れます。一番左Aが０、Bが１です。*/
var titleCol = 1;
var descriptionCol = 2;
var eventDateCol = 3;
var reminder1Col = 4;
var reminder2Col = 5;
var eventURLCol = 6;
var eventIDCol = 7;

/* スプレッドシートのデータ（ヘッダーの行は除く）が入力してある始めのセルと終わりの列を指定します。現在はA2セルからH列まで。*/
var dataRange = "A2:H";

/* 動作の流れ　

1．各行の情報を取得する。
2．イベントURLが空白の場合、新しくイベントをカレンダーに追加し(addEventToCalendar)、イベントURLとIDをシートに入力する(setEventResultToSheet)。
3．イベントURLが入力されている場合、イベントを更新する(updateEvent)。この際、通知を一度すべて削除し、新しく通知を設定する。

4．updateCalendarを、トリガーを設定して実行する。

*/

function updateCalendar() {
  var thisSs = SpreadsheetApp.getActiveSpreadsheet();
  var calendarDataSheet = thisSs.getSheetByName(sheetName);

  var lr = calendarDataSheet.getLastRow();
  var count = calendarDataSheet.getRange(dataRange + lr + "").getValues();
  
  // Check the each row data. If no URL, create a new all day event
  for (i = 0; i < count.length; i++) {
    var row = count[i];

    row.rowNumber = i + 2;

    var title = row[titleCol];
    
    var details = row[descriptionCol];
    var eventDate = row[eventDateCol];
    var reminderMinutes1 = row[reminder1Col]!==0? convertDaysToMinutes(row[reminder1Col]) : 0;
    var reminderMinutes2 = row[reminder2Col]!==0? convertDaysToMinutes(row[reminder2Col]) : 0;
    var eventURL = row[eventURLCol];
    var eventId = row[eventIDCol];
    

    // Logger.log("title=" + title);
    // Logger.log("eventDate original=" + eventDate);
    // Logger.log("description=" + details);
    // Logger.log("eventURL=" + eventURL);
    // Logger.log("eventId=" + eventId);
    // Logger.log("reminderMinutes1=" + reminderMinutes1);
    // Logger.log("reminderMinutes2=" + reminderMinutes2);

    var options = { description: details };

    // Pause a bit to avoid "too many script calling error"
    if (i % 10 == 0) { Utilities.sleep(3000); }

    // Run the following only when title and event date are set.
    if(title!=="" && eventDate!==""){

      // No URL has been set yet = new event. Add this event to the calendar and set the URL and ID in the row.
      if (eventURL == "") {
      
        var newEventResult = addEventToCalendar(title, eventDate, options);
        eventURL = newEventResult[0];
        eventId = newEventResult[1];
        setEventResultToSheet(calendarDataSheet, row.rowNumber, eventURL, eventId) ; 
  
        // Add reminders
        if(reminderMinutes1 > 0){
          addReminder(eventId, reminderMinutes1);
        }
        if(reminderMinutes2 > 0){
          addReminder(eventId, reminderMinutes2);
        }
      }
      // URL is set. Update the event.
      else{
        updateEvent(eventId, title, eventDate, details, reminderMinutes1, reminderMinutes2);
      }
    }
    
  }
}



/* Add a new event to calendar. Return the event URL and ID.
*/
function addEventToCalendar(title, eventDate, options) {

  // Logger.log("addEventToCalendar start");

  var eventCal = CalendarApp.getCalendarById(CALENDAR_ID);
  var newEvent = eventCal.createAllDayEvent(title, eventDate, options);
  var newEventId = newEvent.getId();

  // Logger.log("title=" + title + " newEventId " + newEventId);

  var eventURL = "https://calendar.google.com/calendar/u/0/r/eventedit/" + getEventIdForURL(newEventId);

  // Logger.log("eventURL=" + eventURL);

  // Logger.log("addEventToCalendar end");

  var returnArray = [eventURL, newEventId];

  return returnArray;
}



/* Update an event. Run this every time for every row which has URL.
*/
function updateEvent(eventId, title, date, description, reminderMinutes1, reminderMinutes2) {

  var event = CalendarApp.getCalendarById(CALENDAR_ID).getEventById(eventId);

  // Remove all reminders once.
  event.removeAllReminders();  
  event.setTitle(title);
  event.setAllDayDate(date);
  event.setDescription(description) ;

  if(reminderMinutes1 > 0){
    addReminder(eventId, reminderMinutes1);
  }
  if(reminderMinutes2 > 0){
    addReminder(eventId, reminderMinutes2);
  }
  
  // Logger.log("updateEvent end");
}



/* Set event URL and ID back in the spreadsheet, used after a new event is added to the calendar. 
*/
function setEventResultToSheet(sheet, rowNumber, eventURL, eventId) {

  // Logger.log("setEventURLToSheet start");
  sheet.getRange(rowNumber, eventURLCol+1).setValue(eventURL);  //note - the column number has to be added by 1
  sheet.getRange(rowNumber, eventIDCol+1).setValue(eventId);  //note - the column number has to be added by 1
  // Logger.log("setEventURLToSheet end");
}



/* Add a popup reminder to an event. Call this after a new event created or at updating an existing event. 
*/
function addReminder(eventId, minutesBefore){
  // Logger.log("addReminder start");
  var event = CalendarApp.getCalendarById(CALENDAR_ID).getEventById(eventId);
  event.addPopupReminder(minutesBefore);
  // Logger.log("addReminder end");
}



/* Utility function for getting the new event URL to open the edit page on the browser.
 Please note that this function uses Google Calendar API, which you need to enable the console
 refer https://qiita.com/satoshiks/items/db19d5d15bf376faf083#calendar-api%E3%82%92%E6%9C%89%E5%8A%B9%E3%81%AB%E3%81%99%E3%82%8B%E6%89%8B%E9%A0%86
*/
function getEventIdForURL(eventId) {

  // details of this method can be referred here -> https://developers.google.com/calendar/v3/reference/events/list
  var events = Calendar.Events.list(CALENDAR_ID, {
    iCalUID: eventId,
    singleEvents: true,
    orderBy: "startTime",
    maxResults: 1
  });
  var link = events.items[0].htmlLink;
  var eventIdOnCalendar = link.split("eid=")[1];
  // Logger.log(eventIdOnCalendar);

  return eventIdOnCalendar;
}



/* Utility function for converting days to minutes 
*/
function convertDaysToMinutes(days){

  // (days x 24 hours x 60 minutes)
  return days*24*60; 
}



/* Add menu in the spreadsheet
https://developers.google.com/apps-script/guides/menus
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('カレンダー連携')
      .addItem('実行', 'updateCalendar')
      .addToUi();
}

