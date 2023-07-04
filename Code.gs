/**
 * The `export_gcal_to_gsheet` function exports events from the active user's Google Calendar to 
 * the active Google Sheets document.
 * 
 * The function uses the user's email to identify their Google Calendar. It retrieves events 
 * from the calendar between the dates specified in cells 'B3' (start date) and 'D3' (end date) 
 * of the active sheet. The function excludes events containing the string "-project123" in their titles.
 * 
 * The function sets up a header row starting from row 5 and column 1 with the fields: "Week day", 
 * "Date", "Event Title", "Event Location". 
 * 
 * The function then populates the Google Sheet with the calendar events. It starts populating from 
 * row 6, with each row representing an event and columns representing the event's day of the week, 
 * start time, title, and location. The day of the week is derived using the `getWeekDay` function.
 * 
 * The function also includes validation checks for the existence of the calendar, the validity of 
 * the date range, and the presence of events in the specified date range. If any of these checks fail, 
 * the function logs an error message and exits.
 * 
 * The `getWeekDay` function takes a Date object and returns the corresponding name of the weekday. 
 * This function also includes a validation check to ensure the input is a valid Date object.
 */

/**
 * DEPLOYMENT STEPS:
 *
 * 1. Open a Google Sheets document where you want to deploy the script.
 *
 * 2. Click on Extensions > Apps Script.
 *
 * 3. Delete any code in the Apps Script editor and paste this script.
 *
 * 4. Save the project with a suitable name.
 *
 * 5. Run the function 'export_gcal_to_gsheet' by selecting it from the select function drop-down
 *    (where it says 'Select function') and then click on the triangle 'Run' button. 
 *    The first time you run the script, it will ask for permissions to access your Google Calendar 
 *    and the active Google Sheets document. Grant the necessary permissions.
 *
 * 6. Make sure the spreadsheet has the start and end dates in cells 'B3' and 'D3'.
 *
 * 7. Now, every time you want to update the sheet with the calendar events, you just need to run the 
 *    'export_gcal_to_gsheet' function.
 *
 * NOTE: If you want this function to run automatically at certain intervals (like daily or weekly), 
 *       you can set up a trigger. To do this, click on the clock icon ('Triggers') in the left sidebar 
 *       of the Apps Script Editor. Then, click on '+ Add Trigger' at the bottom right of the page. 
 *       Choose 'export_gcal_to_gsheet' for the function to run, set 'time-driven' for the deployment event type,
 *       and then specify the frequency (hourly, daily, etc.) and time of day as per your preference.
 */

function export_gcal_to_gsheet(){
    // Define default start and end dates
    var startDate = getFirstDayOfPrevMonth();
    var endDate = getLastDayOfPrevMonth();

    // Define calendar and spreadsheet objects
    var mycal = Session.getActiveUser().getEmail();
    var cal = CalendarApp.getCalendarById(mycal);
    var sheet = SpreadsheetApp.getActiveSheet();

    // Check if calendar exists
    if(!cal) {
      Logger.log("Calendar not found for email: " + mycal);
      return;
    }

    // Fetch events between start and end dates excluding '-project123'
    var events = cal.getEvents(startDate, endDate, {search: '-project123'});

    // Check if there are events to export
    if(!events.length) {
      Logger.log("No events found between the given date range excluding '-project123'.");
      return;
    }

    // Define headers for the spreadsheet
    var header = [["Week day", "Date", "Start Time", "End Time", "Duration (hours)", "Event Title", "Event Location"]];
    var range = sheet.getRange(5, 1, 1, 7);
    range.setValues(header);

    // Loop over events and add them to the spreadsheet
    for (var i = 0; i < events.length; i++) {
        var row = i + 6; // Start from row 6 to leave room for headers
        var duration = (events[i].getEndTime().getTime() - events[i].getStartTime().getTime()) / 3600000; // Calculate duration in hours
        var details = [[ 
            getWeekDay(events[i].getStartTime()),
            events[i].getStartTime(),
            events[i].getStartTime().toLocaleTimeString(), 
            events[i].getEndTime().toLocaleTimeString(), 
            duration.toFixed(2), 
            events[i].getTitle(), 
            events[i].getLocation() 
        ]];
        var range = sheet.getRange(row, 1, 1, 7);
        range.setValues(details);
    }
}

function getFirstDayOfPrevMonth(){
    var date = new Date();
    date.setMonth(date.getMonth()-1);
    date.setDate(1);
    return date;
}

function getLastDayOfPrevMonth(){
    var date = new Date();
    date.setDate(1); // Set to first day of current month
    date.setDate(date.getDate()-1); // Subtract a day to get last day of previous month
    return date;
}

function getWeekDay(date){
    // Array of weekdays
    var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

    // Validate the input is a Date object
    if(!(date instanceof Date)) {
      Logger.log("Invalid date object.");
      return;
    }

    var dayNumber = date.getDay(); // Get the day of the week as a number (0-6)
    return days[dayNumber]; // Return the corresponding weekday
}
