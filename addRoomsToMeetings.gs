/**
 * Google Apps Script to check upcoming calendar meetings and assign rooms.
 *
 * This script looks for meetings:
 * - Owned by the user running the script (for the automated check).
 * - On days when the user is marked as "in office" (defined in inOfficeDays function).
 * - That do not currently have a meeting room assigned - from the AVAILABLE_ROOMS list.
 *
 * It then attempts to find and assign an available meeting room within the specified office location and floor.
 *
 * This version also includes functionality to assign rooms to meetings not owned by the user as:
 * - Populate a Google Sheet with upcoming meeting details.
 * - Allow users to manually select meetings in the sheet for room assignment.
 * - Assign rooms to selected meetings from the sheet.
 */

// --- Configuration ---
// Note: MY_OFFICE_LOCATION and MY_OFFICE_FLOOR are for conceptual use or future expansion
// They are not directly used in the current room assignment logic based on email addresses.
// const MY_OFFICE_LOCATION = "Bellevue Office";
// const MY_OFFICE_FLOOR = "4th Floor";

const CALENDAR_ID = "primary"; // Use "primary" for your main calendar
const DAYS_TO_CHECK = 30;     // Check for meetings in the next 30 days for auto-assignment and sheet population

const SHEET_NAME = "CALENDAR";                 // The name of the sheet tab you want to use

// Define your in-office days here.  Modify this array to match your schedule.
const IN_OFFICE_DAYS = [1, 2, 3, 4]; // Monday (1), Wednesday (3), Thursday (4)

// Define your available meeting rooms. These are typically email addresses like "room-name@yourdomain.com".
// You'll need to get the exact resource IDs for your rooms.
// Ensure these are lowercase for consistent comparison.
const AVAILABLE_ROOMS = {
    "resource_id@resource.calendar.google.com": "Resource Display Name",
   
};

/**
 * Creates a custom menu in the Google Sheet when it is opened.
 */

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    /**   
      var roomSubmemu = ui.createMenu('Room')
          .addItem('Assign Rooms to Selected Meetings', 'assignRoomsFromSheet')     // New: Assign from sheet
          .addItem('Assign Rooms to Owned Meetings', 'assignMeetingRooms') // Existing auto-assignment
    
      var timeSubmenu = ui.createMenu('Time')
          .addItem('Shift Selected Meetings by 5 Mins', 'shiftSelectedMeetings') // NEW: Shift meetings
          .addItem('Shift Owned Meetings by 5 Minutes','updateOwnedCalendarEntries' ) // Move owned meetings by 5 minutes. 
    
    */
    ui.createMenu('Calendar Tools') // Name of your custom menu
        .addItem('Populate Meetings to Sheet', 'populateMeetingsToSheet') // New: Populate sheet
        .addSeparator()
        //      .addSubMenu(roomSubmemu)
        .addItem('Assign Rooms to Selected Meetings', 'assignRoomsFromSheet')     // New: Assign from sheet
        .addItem('Shift Selected Meetings by 5 Mins', 'shiftSelectedMeetings') // NEW: Shift meetings
        .addSeparator()
        //      .addSubMenu(timeSubmenu)
        .addItem('Assign Rooms to Owned Meetings', 'assignMeetingRooms') // Existing auto-assignment
        .addItem('Shift Owned Meetings by 5 Minutes', 'updateOwnedCalendarEntries') // Move owned meetings by 5 minutes. 
        .addSeparator()
        .addToUi();
}


/**
 * Assigns Meeting Rooms automatically to the meetings in next "DAYS_TO_CHECK" that:
 * - Is owned by the user
 * - User is in Office
 * - Doesn't have a meeting room assigned from the above list.
 * Any meeting room assigned that is not in above list would be ignored.
 *
 */
// --- Main Auto-Assignment Function ---
function assignMeetingRooms() {
    Logger.log("Starting automated meeting room assignment script...");

    const now = new Date();
    const oneMonthLater = new Date();
    oneMonthLater.setDate(now.getDate() + DAYS_TO_CHECK);

    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
        Logger.log(`Calendar with ID "${CALENDAR_ID}" not found.`);
        return;
    }

    // Get events for the next month
    const events = calendar.getEvents(now, oneMonthLater);
    Logger.log(`Found ${events.length} events in the next ${DAYS_TO_CHECK} days.`);

    events.forEach(event => {
        try {
            processEvent(event);
        } catch (e) {
            Logger.log(`Error processing event '${event.getTitle()}': ${e.message}`);
        }
    });

    Logger.log("Automated meeting room assignment script finished.");
}

/**
 * Populates a Google Sheet with upcoming meeting details,
 * including a checkbox column for room assignment and event ID.
 */
function populateMeetingsToSheet() {
    Logger.log("Populating meetings to sheet...");
    try {
        //const ss = SpreadsheetApp.openById(SHEET_ID);
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) {
            Logger.log(`Sheet "${SHEET_NAME}" not found in spreadsheet. Please check configuration.`);
            SpreadsheetApp.getUi().alert('Error', `Sheet "${SHEET_NAME}" not found. Please verify SHEET_NAME in the script configuration.`, SpreadsheetApp.getUi().ButtonSet.OK);
            return;
        }

        // Clear existing data but keep header
        sheet.clearContents();

        const headers = ["Select", "Date", "Start Time", "End Time", "Title", "Organizer", "Current Room", "Event ID"];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

        const now = new Date();
        const futureDate = new Date();
        futureDate.setDate(now.getDate() + DAYS_TO_CHECK);

        const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
        if (!calendar) {
            Logger.log(`Calendar with ID "${CALENDAR_ID}" not found.`);
            SpreadsheetApp.getUi().alert('Error', `Calendar "${CALENDAR_ID}" not found. Please verify CALENDAR_ID.`, SpreadsheetApp.getUi().ButtonSet.OK);
            return;
        }

        const events = calendar.getEvents(now, futureDate);
        const data = [];

        events.forEach(event => {
            // Skip all-day events as rooms are usually for specific times
            if (event.isAllDayEvent()) {
                return;
            }

            const eventId = event.getId();
            const eventTitle = event.getTitle();
            const eventStartTime = event.getStartTime();
            const eventEndTime = event.getEndTime();
            const eventOrganizer = event.getCreators().length > 0 ? event.getCreators()[0] : "N/A";
            const hasRoomAlready = hasMeetingRoom(event) ? "Yes" : "No"; // Use the correct hasMeetingRoom logic

            // Row structure: [Checkbox, Date, Start Time, End Time, Title, Organizer, Current Room, Event ID]
            data.push([
                false, // Default checkbox value
                Utilities.formatDate(eventStartTime, Session.getScriptTimeZone(), "yyyy-MM-dd"),
                Utilities.formatDate(eventStartTime, Session.getScriptTimeZone(), "HH:mm"),
                Utilities.formatDate(eventEndTime, Session.getScriptTimeZone(), "HH:mm"),
                eventTitle,
                eventOrganizer,
                hasRoomAlready,
                eventId // Hidden Event ID
            ]);
        });

        if (data.length > 0) {
            sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
        }

        // Set "Select" column as checkboxes
        const checkboxRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
        checkboxRange.insertCheckboxes();

        // Hide the "Event ID" column (Column H)
        sheet.hideColumns(8); // Column H is the 8th column

        sheet.autoResizeColumns(1, headers.length - 1); // Auto-resize visible columns
        SpreadsheetApp.getUi().alert('Success', `Successfully populated ${data.length} meetings to the sheet.`, SpreadsheetApp.getUi().ButtonSet.OK);
        Logger.log("Meetings populated successfully.");

    } catch (e) {
        Logger.log(`Error populating meetings to sheet: ${e.message}`);
        SpreadsheetApp.getUi().alert('Error', `Failed to populate meetings: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * Reads selected meetings from the Google Sheet and attempts to assign rooms.
 */
function assignRoomsFromSheet() {
    Logger.log("Assigning rooms from sheet...");
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) {
            Logger.log(`Sheet "${SHEET_NAME}" not found in spreadsheet. Please check configuration.`);
            SpreadsheetApp.getUi().alert('Error', `Sheet "${SHEET_NAME}" not found. Please verify SHEET_NAME in the script configuration.`, SpreadsheetApp.getUi().ButtonSet.OK);
            return;
        }

        // Get all data starting from the second row (skipping headers)
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
            SpreadsheetApp.getUi().alert('Info', 'No meetings found in the sheet to process.', SpreadsheetApp.getUi().ButtonSet.OK);
            Logger.log("No meetings found in the sheet to process.");
            return;
        }

        const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
        const values = dataRange.getValues();
        const updateStatuses = []; // To store updates for the "Current Room" column

        let roomsAssignedCount = 0;

        for (let i = 0; i < values.length; i++) {
            const row = values[i];
            const assignRoomCheckbox = row[0]; // Column A: "Assign Room?" (TRUE/FALSE)
            const eventId = row[7];           // Column H: "Event ID" (hidden)

            if (assignRoomCheckbox === true) { // Check if the checkbox is ticked
                Logger.log(`Attempting to assign room for event ID: ${eventId}`);
                try {
                    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
                    const event = calendar.getEventById(eventId);

                    if (event) {
                        // Re-check if it has a room to prevent re-assigning if manually added elsewhere
                        if (hasMeetingRoom(event)) {
                            Logger.log(`Event '${event.getTitle()}' (ID: ${eventId}) already has a room. Skipping.`);
                            updateStatuses.push(["Already assigned"]); // Update sheet status
                            continue;
                        }

                        const assigned = assignAvailableRoom(event, event.getStartTime(), event.getEndTime());
                        if (assigned) {
                            updateStatuses.push([`Assigned: ${getAssignedRoomEmail(event) || 'N/A'}`]); // Get actual assigned room
                            roomsAssignedCount++;
                        } else {
                            updateStatuses.push(["No room found"]);
                        }
                    } else {
                        Logger.log(`Event with ID ${eventId} not found in calendar. It might have been deleted.`);
                        updateStatuses.push(["Event not found"]);
                    }
                } catch (e) {
                    Logger.log(`Error processing row ${i + 2} (Event ID: ${eventId}): ${e.message}`);
                    updateStatuses.push([`Error: ${e.message}`]);
                }
            } else {
                updateStatuses.push([row[6]]); // Keep existing status if not selected for assignment
            }
        }

        // Update the "Current Room" column in the sheet
        if (updateStatuses.length > 0) {
            sheet.getRange(2, 7, updateStatuses.length, 1).setValues(updateStatuses); // Column G is the 7th column
        }

        SpreadsheetApp.getUi().alert('Success', `Finished processing selected meetings. Rooms assigned: ${roomsAssignedCount}`, SpreadsheetApp.getUi().ButtonSet.OK);
        Logger.log("Room assignment from sheet finished.");

    } catch (e) {
        Logger.log(`Error assigning rooms from sheet: ${e.message}`);
        SpreadsheetApp.getUi().alert('Error', `Failed to assign rooms from sheet: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}




/**
 * Processes a single calendar event for automated assignment based on user ownership and in-office days.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The calendar event to process.
 */
function processEvent(event) {
    const eventTitle = event.getTitle();
    const eventCreators = event.getCreators();
    const eventOwnerEmail = eventCreators && eventCreators.length > 0 ? eventCreators[0] : null;
    const eventStartTime = event.getStartTime();
    const eventEndTime = event.getEndTime();

    // 1. Check if owned by me (the script runner)
    const scriptOwnerEmail = Session.getActiveUser().getEmail();
    const isOwnedByMe = eventOwnerEmail && eventOwnerEmail.toLowerCase() === scriptOwnerEmail.toLowerCase();

    if (!isOwnedByMe) {
        // Logger.log(`Skipping event '${eventTitle}': Not owned by me.`);
        return;
    }

    // Skip all-day events as rooms are usually for specific times
    if (event.isAllDayEvent()) {
        Logger.log(`Skipping event '${eventTitle}': is an all-day event.`);
        return;
    }

    // 2. Check if on a day I am in office
    if (!isInOfficeDay(eventStartTime)) {
        // Logger.log(`Skipping event '${eventTitle}': Not an in-office day.`);
        return;
    }

    // 3. Check if it already has a meeting room assigned
    if (hasMeetingRoom(event)) {
        //Logger.log(`Skipping event '${eventTitle}': Already has a meeting room.`);
        return;
    }

    Logger.log(`Processing eligible event for auto-assignment: '${eventTitle}' on ${eventStartTime.toLocaleString()}`);

    // Attempt to find and assign a room
    const assigned = assignAvailableRoom(event, eventStartTime, eventEndTime);
    if (assigned) {
        Logger.log(`Successfully auto-assigned a room to: '${eventTitle}'`);
    } else {
        Logger.log(`Could not find an available room for auto-assignment for: '${eventTitle}'`);
    }
}

/**
 * Checks if the given date is an "in-office" day based on a specific list of days.
 * @param {Date} date The date to check.
 * @returns {boolean} True if the date is an in-office day, false otherwise.
 */
function isInOfficeDay(date) {
    const dayOfWeek = date.getDay(); // Sunday - 0, Monday - 1, ..., Saturday - 6

    return IN_OFFICE_DAYS.includes(dayOfWeek);
}

/**
 * Checks if an event already has a meeting room assigned.
 * It does this by getting the list of all guest email addresses and checking if any
 * of them match an email in the predefined AVAILABLE_ROOMS list.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The calendar event.
 * @returns {boolean} True if a meeting room is assigned, false otherwise.
 */
function hasMeetingRoom(event) {
    // getGuestList() returns an array of email strings for all guests.
    const guestEmails = event.getGuestList();
    //const guestEmails = event.getGuests();
    for (let i = 0; i < guestEmails.length; i++) {
        //  const guestEmail = guestEmails[i];
        const guestEmail = guestEmails[i].getEmail().toLowerCase();
        //Logger.log("Type of guestEmail: " + typeof guestEmail + ", Value: " + guestEmail);      
        //const guestEmail = guestEmails[i]; // Convert to lowercase for case-insensitive comparison
        if (AVAILABLE_ROOMS.hasOwnProperty(guestEmail)) { // Check if email is a key in the object
            Logger.log(`Skipping Event: '${event.getTitle()}' already has room: ${AVAILABLE_ROOMS[guestEmail]}`); // Log the friendly name
            return true;
        }
    }
    return false;
}

/**
 * Finds and assigns an available meeting room to an event.
 * Checks room availability by getting the room's calendar and querying its events.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The calendar event to update.
 * @param {Date} startTime The start time of the event.
 * @param {Date} endTime The end time of the event.
 * @returns {boolean} True if an available room was assigned, false otherwise.
 */
function assignAvailableRoom(event, startTime, endTime) {
    // for (const roomEmail of AVAILABLE_ROOMS) {
    // Iterate through the keys (resource IDs) of the AVAILABLE_ROOMS object
    for (const roomEmail of Object.keys(AVAILABLE_ROOMS)) {

        const roomName = AVAILABLE_ROOMS[roomEmail]; // Get the friendly name for logging

        try {
            // 1. Get the Calendar object for the room using its email as the ID.
            const roomCalendar = CalendarApp.getCalendarById(roomEmail);

            // Ensure the room email corresponds to a valid calendar that we can access.
            if (!roomCalendar) {
                Logger.log(`Could not retrieve calendar for room: ${roomName} (${roomEmail}). It might not be a valid resource email or permissions are missing.`);
                continue; // Try the next room
            }

            // 2. Check if the room calendar has any events overlapping with the meeting time.
            const existingEventsForRoom = roomCalendar.getEvents(startTime, endTime);

            if (existingEventsForRoom.length === 0) {
                // If no existing events are found, the room is available.
                // 3. Attempt to add the room as a guest to the event.
                event.addGuest(roomEmail);
                Logger.log(`Successfully assigned available room ${roomName} to event '${event.getTitle()}'`);
                return true; // Room found and assigned
            } else {
                // Room is occupied
                // Logger.log(`Room ${roomEmail} is occupied for event '${event.getTitle()}' from ${startTime.toLocaleString()} to ${endTime.toLocaleString()}.`); // Too verbose
                // Continue to the next room
            }
        } catch (e) {
            // This catch block handles potential errors during the process, such as:
            // - Problems getting the calendar by ID (e.g., roomEmail is not a valid calendar ID).
            // - Errors when adding the guest (e.g., transient network issues, permissions on the event itself).
            Logger.log(`Error when processing room ${roomName} for event '${event.getTitle()}': ${e.message}`);
            continue; // Try the next room
        }
    }
    return false; // No available room could be assigned after trying all
}

/**
 * Helper function to get the email of an assigned room from an event.
 * Used for updating the sheet with the specific room assigned.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The event to check.
 * @returns {string|null} The email of the first assigned room found, or null if none.
 */
function getAssignedRoomEmail(event) {
    const guestEmails = event.getGuestList();
    //const guestEmails = event.getGuests();

    for (const guestEmail of guestEmails) {
        //.     Logger.log("Type of guestEmail: " + typeof guestEmail + ", Value: " + guestEmail);
        //const lowerCaseGuestEmail = guestEmail.toLowerCase();
        const lowerCaseGuestEmail = guestEmail.getEmail().toLowerCase();
        if (AVAILABLE_ROOMS.hasOwnProperty(lowerCaseGuestEmail)) { // Check if email is a key in the object
            return AVAILABLE_ROOMS[lowerCaseGuestEmail]; // Return the friendly name
        }

    }
    return null;
}

// --- Setup Trigger Function (Run this once to set up the daily trigger) ---
/**
 * Sets up a daily time-driven trigger to run the 'assignMeetingRooms' function.
 * This function should be run once manually to establish the trigger.
 */
function setupTrigger() {
    // Delete existing triggers to avoid duplicates
    ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getHandlerFunction() === 'assignMeetingRooms') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // Create a new time-driven trigger to run daily between 1 AM and 2 AM
    ScriptApp.newTrigger('assignMeetingRooms')
        .timeBased()
        .everyDays(1)
        .atHour(1) // Run between 1:00 AM and 2:00 AM
        .create();

    Logger.log("Daily trigger for 'assignMeetingRooms' has been set up.");
    SpreadsheetApp.getUi().alert('Trigger Setup', 'Daily trigger for automated room assignment has been set up (runs between 1-2 AM).', SpreadsheetApp.getUi().ButtonSet.OK);
}