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
function shiftOwnedCalendarEntries() {

    //const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    //const daysCell = sheet.getRange('B1').getValue(); // This if you want to read the number of days from sheet
    const daysCell = DAYS_TO_CHECK; // Change to how many days in future to look
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


// --- NEW Function: Shift Selected Meetings ---
/**
 * Shifts the start and end times of selected meetings in the sheet by 5 minutes later,
 * if they start on the hour or half-hour AND their checkbox is selected.
 * Updates both the Calendar event and the Google Sheet.
 * 
 * Note: Use this in conjunction with addRoomsToMeetings.gs --> populateMeetingsToSheet() to fetch the meetings to the sheet. 
 */
function shiftSelectedMeetings() {
    Logger.log("Starting to shift selected meetings...");
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) {
            Logger.log(`Sheet "${SHEET_NAME}" not found.`);
            SpreadsheetApp.getUi().alert('Error', `Sheet "${SHEET_NAME}" not found. Please verify SHEET_NAME.`, SpreadsheetApp.getUi().ButtonSet.OK);
            return;
        }

        const selectedRange = sheet.getActiveRange();
        if (!selectedRange || selectedRange.getRow() === 1) { // Check if a range is selected and it's not just the header
            SpreadsheetApp.getUi().alert('Info', 'Please select the rows containing the meetings you wish to shift (excluding the header).', SpreadsheetApp.getUi().ButtonSet.OK);
            Logger.log("No valid range selected for shifting meetings.");
            return;
        }

        const startRow = 2;
        const endRow = sheet.getLastRow();
        const numRows = endRow - startRow + 1;
        if (endRow < 2) {
            SpreadsheetApp.getUi().alert('Info', 'No meetings found in the sheet to process.', SpreadsheetApp.getUi().ButtonSet.OK);
            Logger.log("No meetings found in the sheet to process.");
            return;
        }

        // Get all relevant data for the selected rows (from Column A to H)
        // This ensures we always have the checkbox and Event ID at their known positions
        const dataToProcess = sheet.getRange(startRow, 1, numRows, 8).getValues(); // Get from Column A to H (8 columns)

        const updatedTimesForSheet = []; // To store updates for Start Time (Col C) and End Time (Col D)
        let shiftedCount = 0;

        for (let i = 0; i < dataToProcess.length; i++) {
            const rowData = dataToProcess[i];
            const assignCheckbox = rowData[0]; // Column A (index 0)
            const eventId = rowData[7];        // Column H (index 7)
            const currentSheetDate = rowData[1]; // Column B (index 1)
            const currentSheetStartTimeValue = rowData[2]; // Column C (index 2)
            const currentSheetEndTimeValue = rowData[3];   // Column D (index 3)

            // Only process if the checkbox is true AND an Event ID exists
            if (assignCheckbox === true && eventId) {
                // Convert sheet date/time values to Date objects for manipulation
                // Date objects from getValues() for date/time columns might be Date objects directly
                // or string representations depending on formatting. It's safer to reconstruct.
                const eventDate = new Date(currentSheetDate);
                const startTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(),
                    currentSheetStartTimeValue.getHours(), currentSheetStartTimeValue.getMinutes(), 0);
                const endTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(),
                    currentSheetEndTimeValue.getHours(), currentSheetEndTimeValue.getMinutes(), 0);

                const minutes = startTime.getMinutes();

                if (minutes === 0 || minutes === 30) {
                    Logger.log(`Processing event '${rowData[4]}' (ID: ${eventId}) starting at ${startTime.toLocaleString()}`);

                    try {
                        const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
                        const event = calendar.getEventById(eventId);

                        if (event) {
                            const newStartTime = new Date(startTime.getTime() + 5 * 60 * 1000); // Add 5 minutes
                            const newEndTime = new Date(endTime.getTime() + 5 * 60 * 1000);   // Add 5 minutes

                            event.setTime(newStartTime, newEndTime); // Update event in Google Calendar
                            Logger.log(`Shifted event '${event.getTitle()}' from ${startTime.toLocaleTimeString()} to ${newStartTime.toLocaleTimeString()}`);
                            shiftedCount++;

                            // Prepare for sheet update for this specific row
                            const newStartTimeFormatted = Utilities.formatDate(newStartTime, Session.getScriptTimeZone(), "HH:mm");
                            const newEndTimeFormatted = Utilities.formatDate(newEndTime, Session.getScriptTimeZone(), "HH:mm");
                            updatedTimesForSheet.push({
                                rowIdx: i, // Index within the `dataToProcess` array
                                newST: newStartTimeFormatted,
                                newET: newEndTimeFormatted
                            });

                        } else {
                            Logger.log(`Event with ID ${eventId} not found in calendar for shifting. It might have been deleted.`);
                        }
                    } catch (e) {
                        Logger.log(`Error shifting event (ID: ${eventId}): ${e.message}`);
                    }
                } else {
                    // Logger.log(`Skipping event '${rowData[4]}': does not start on the hour or half-hour.`);
                }
            } else {
                // Logger.log(`Skipping row ${startRow + i} because checkbox not true or no Event ID.`);
            }
        }

        // Perform batch update on the sheet for modified rows
        if (updatedTimesForSheet.length > 0) {
            for (const updateInfo of updatedTimesForSheet) {
                // Calculate the actual row number in the sheet
                const actualSheetRow = startRow + updateInfo.rowIdx;
                sheet.getRange(actualSheetRow, 3).setValue(updateInfo.newST); // Column C (Start Time)
                sheet.getRange(actualSheetRow, 4).setValue(updateInfo.newET); // Column D (End Time)
            }
        }

        SpreadsheetApp.getUi().alert('Success', `Finished shifting meetings. Total meetings shifted: ${shiftedCount}.`, SpreadsheetApp.getUi().ButtonSet.OK);
        Logger.log("Finished shifting selected meetings.");

    } catch (e) {
        Logger.log(`An unexpected error occurred while shifting meetings: ${e.message}`);
        SpreadsheetApp.getUi().alert('Error', `Failed to shift meetings: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}
