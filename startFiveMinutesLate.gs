/**
 * Google Apps Script that checks for meetings starting at the hour or 30 minutes and move them by 5 minutes.  
 * The intent of this script is to give you enough time to move between the rooms and meeting overruns, rooms taking time to free up.
 * You want to turn on speedy meetings for this so that the meetings finish off at the correct time.
 * 
 * This script looks for meetings up to 30 days in future. :
 * - Owned by the user running the script.
 * - Are starting at the hour or 30 minutes past hour
 * 
 *
 * It then moves the start and end time by 5 minutes. 
 */
// Function 2: Update owned events that start at HH:00 or HH:30 to start 5 minutes later
function updateOwnedCalendarEntries() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //const daysCell = sheet.getRange('B1').getValue(); // This if you want to read the number of days from sheet
  const daysCell = 30; // Change to how many days in future to look
  const calendar = CalendarApp.getCalendarById('primary');
  const now = new Date();
  const end = new Date(now.getTime() + daysCell * 24 * 60 * 60 * 1000);
  const userEmail = Session.getActiveUser().getEmail();

  const events = calendar.getEvents(now, end);
  for (const event of events) {
    let owner = "";
    try {
      owner = typeof event.getCreators === "function"
        ? (event.getCreators().length > 0 ? event.getCreators()[0] : "")
        : (typeof event.getCreator === "function" ? event.getCreator() : "");
    } catch (e) { }

    // Check if event is owned by user
    const isOwnedByUser = (owner === userEmail || owner === "" || (typeof event.isOwnedByMe === "function" && event.isOwnedByMe()));
    const start = event.getStartTime();
    const endTime = event.getEndTime();
    const startMinute = start.getMinutes();
    if (event.isAllDayEvent()) {
      Logger.log("Event: " + event.getTitle() + " is all day event:")
      continue; // Skip all-day events
    }
    if (isOwnedByUser && (startMinute === 0 || startMinute === 30)) {
      // Shift meeting start and end times by 5 minutes later
      Logger.log("Event: " + event.getTitle() + " is going to be modified:")
      const newStart = new Date(start.getTime() + 5 * 60000);
      const newEnd = new Date(endTime.getTime() + 5 * 60000);
      Logger.log("Event modified to New Start of:" + newStart + "and End time of:" + newEnd);
      event.setTime(newStart, newEnd);
    }
  }
}



/**
 * Google Apps Script to fetch calendar events and log them to the active sheet. Needs cell B1 to tell how many days in future to look. 
 * 
 * 
 */

function fetchCalendarEvents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const daysCell = sheet.getRange('B1').getValue(); // configurable number of days
  const calendar = CalendarApp.getCalendarById('primary'); // Adjust Calendar ID if needed
  const now = new Date();
  const end = new Date(now.getTime() + daysCell * 24 * 60 * 60 * 1000);

  // Optional: clear previous data except config/header
  sheet.getRange("A3:F").clearContent();

  // Header row (row 2)
  sheet.getRange('A2:F2').setValues([["Summary", "Description", "Start Date/Time",
    "End Date/Time", "Owner", "Attendee List"]]);

  const events = calendar.getEvents(now, end);
  let output = [];
  for (const event of events) {
    const summary = event.getTitle();
    const description = event.getDescription();
    const start = event.getStartTime();
    const endTime = event.getEndTime();
    let owner = "";
    try {
      owner = typeof event.getCreators === "function"
        ? (event.getCreators().length > 0 ? event.getCreators()[0] : "")
        : (typeof event.getCreator === "function" ? event.getCreator() : "");
    } catch (e) { }

    let attendees = "";
    try {
      attendees = event.getGuestList().map(g => g.getEmail()).join(", ");
    } catch (e) { }

    output.push([summary, description, start, endTime, owner, attendees]);
  }

  // Write event data to the sheet, starting at row 3
  if (output.length > 0) {
    sheet.getRange(3, 1, output.length, 6).setValues(output);
  }
}
