/*************************************************
*
* GOOGLE CALENDAR DATA AGGREGATOR
* Author: Nick Young
* Email: usaussie@gmail.com
*
* INSTRUCTIONS: 
* (1) Update the variables below to point to your google sheet
* (2) Run the set_sheet_headers() function once, which will prompt for permissions and set the sheet column headers
* (3) Accept the permissions (asking for access for your script to read/write to google drive etc)
* (4) Run the get_yesterdays_events() function (once, or set a trigger to run every day for ongoing aggregation)
* (5) Look in your google sheet as the function is running, and you should see results being inserted
*
* RECOMMENDATION:
* If you run this using a triggered schedule, it will look at events from "yesterday"
* This is then easy to add as a source to a Data Studio Dashboard. 
*
* You can run a backfill by using the variables below, and running the backfill() function.
*
* NOTE: During longer backfills, you may see "CAUGHT EXCEPTION:Error: Could not locate target object while 
*.      calling method getEventsForDay on object with id ___." 
*.      This seems to happen intermittently, and fixes itself if you run it for that same day again. 
*.      So...not sure what's going on there.
*
*
* UPDATE THESE VARIABLES
*
*************************************************/
// google sheet to store aggregate CSV info
var SHEET_URL = "https://docs.google.com/spreadsheets/d/your-sheet-id/edit";
var SHEET_TAB_NAME = "data";

// only needed if you are doing a backfill
var BACKFILL_START_DAYS_AGO = 1; // start X number of days ago for the backfill, and then use the next variable to find out how far
var BACKFILL_NUMBER_OF_DAYS_TO_RETRIEVE = 365; // 365 = grab 365 days worth of calendar events


/*************************************************
*
* DO NOT CHANGE ANYTHING BELOW THIS LINE
*
*************************************************/

/*
*
* ONLY RUN THIS ONCE TO SET THE HEADER ROWS FOR THE GOOGLE SHEETS
*/
function set_sheet_headers() {
  
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(SHEET_TAB_NAME);
  sheet.appendRow(["Event_Id","Creator_Email","Name","Start","End","Location","Guest_Email","Guest_Status","Total_Guest_Count"]);
  
}

function get_yesterdays_events() {
  
  Logger.log("Get Yesterday: " + subDaysFromDate(new Date(),1));
  
  get_events_for_date(subDaysFromDate(new Date(),1));
}

function backfill() {

  var daysArray = getDaysArray(BACKFILL_START_DAYS_AGO, BACKFILL_NUMBER_OF_DAYS_TO_RETRIEVE);
  
  daysArray.forEach(get_events_for_date);
  
}

function get_events_for_date(dateToGet, index, array) {

  var ss = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(SHEET_TAB_NAME);
  
  // Determines how many events are happening on the given day.
  
  try {
    var events = CalendarApp.getDefaultCalendar().getEventsForDay(dateToGet);
  } catch (e) {
    Logger.log('Error get_events_for_date() | Date To Get: ' + dateToGet + ' | CAUGHT EXCEPTION:' + e);
    return;
  }
  
  
  //Logger.log('Get Date: ' + dateToGet);
  Logger.log('Number of events: ' + events.length + " | DateToGet: " + dateToGet);

  if (events[0]) {

    var rowsToWrite = [];

    for (var i=0;i<events.length;i++) {

      // is this an all day event
      if(events[i].isAllDayEvent() == true) {
        //Logger.log('All Day Event: ' + events[i].isAllDayEvent());
        var EVENT_START_TIME = events[i].getAllDayStartDate();
        var EVENT_END_TIME = events[i].getAllDayEndDate();
      } else {
        var EVENT_START_TIME = events[i].getStartTime();
        var EVENT_END_TIME = events[i].getEndTime(); 
      }
      
      //get Guest List
      var guest_list = events[i].getGuestList(true); // get events and include owner

      if (guest_list[0]) {
      
        for (var n=0;n<guest_list.length;n++) {
        
          var newRow = [

            events[i].getId(), 
            events[i].getCreators(),
            events[i].getTitle(), 
            EVENT_START_TIME,            
            EVENT_END_TIME,
            events[i].getLocation(),
            guest_list[n].getEmail(),
            guest_list[n].getGuestStatus(),
            guest_list.length
          ];

          // add to row array instead of append because append is SLOOOOOWWWWW
          rowsToWrite.push(newRow);
        
        }

      } else {

        var newRow = [

          events[i].getId(), 
          events[i].getCreators(),
          events[i].getTitle(), 
          EVENT_START_TIME,            
          EVENT_END_TIME,
          events[i].getLocation(),
          "",
          "",
          1
        ];

        // add to row array instead of append because append is SLOOOOOWWWWW
        rowsToWrite.push(newRow);

      }

    }

    ss.getRange(ss.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite); 

  }
}

function subDaysFromDate(date,d){
  // d = number of day ro substract and date = start date
  var result = new Date(date.getTime()-d*(24*3600*1000));
  return result
}

function getDaysArray(numDaysAgoToStart, daysToGoBack) {
    
    var returnArray = [];
    
    var i = 0;

    while (i < daysToGoBack) {
    
      // Logger.log("Array Loop: " + subDaysFromDate(new Date(),numDaysAgoToStart + i));
      
      returnArray.push(subDaysFromDate(new Date(),numDaysAgoToStart + i));
      i++;

    }

    return returnArray;
    
};


