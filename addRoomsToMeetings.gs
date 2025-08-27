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


// --- Configuration ---
const CALENDAR_ID = "primary"; // Use "primary" for your main calendar.
const DAYS_TO_CHECK = 30;     // Check for meetings in the next 30 days.
const SHEET_NAME = "CALENDAR"; // The name of the sheet tab to use.

// Define your in-office days: Sunday (0), Monday (1), Tuesday (2), etc.
const IN_OFFICE_DAYS = [1, 2, 3, 4]; // Monday, Tuesday, Wednesday, Thursday

// Set the minimum number of participants required to book a room.
const MINIMUM_PARTICIPANTS = 2;

// Set to 'true' to automatically assign rooms for meetings you don't own.
const AUTO_ASSIGN_NON_OWNED_MEETINGS = true;

// Set the maximum number of participants for auto-assigning rooms for non-owned meetings. It treats a group in the invite as one participant.
// A meeting where participants are hidden is treated as one participant.
const NON_OWNED_MEETING_MAX_PARTICIPANTS = 5;

// Define your available meeting rooms. These are typically email addresses like "room-name@yourdomain.com".
// You'll need to get the exact resource IDs for your rooms.
// Ensure these are lowercase for consistent comparison.
const AVAILABLE_ROOMS = {
    "resource_id@resource.calendar.google.com": "Resource Display Name",
   
};

/**
 * Creates a custom menu in the Google Sheet when it is opened.
 */

// --- Menu Setup ---

/**
 * Creates a custom menu in the Google Sheet when it is opened.
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Calendar Tools')
        .addItem('1. Populate Meetings to Sheet', 'populateMeetingsToSheet')
        .addSeparator()
        .addItem('2. Assign Rooms to Selected Meetings', 'assignRoomsFromSheet')
        .addItem('3. Shift Selected Meetings by 5 Mins', 'shiftSelectedMeetings')
        .addSeparator()
        .addItem('Auto-Assign All My Meeting Rooms', 'assignMeetingRooms')
        .addItem('Auto-Shift All My Meetings by 5 Mins', 'shiftOwnedMeetingsBy5Mins')
        .addSeparator()
        .addItem('Setup Daily Trigger', 'setupTrigger')
        .addToUi();
}


// --- Sheet Interaction Functions ---

/**
 * Populates the Google Sheet with upcoming meeting details.
 */
function populateMeetingsToSheet() {
    Logger.log("Populating meetings to sheet...");
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) {
            throw new Error(`Sheet "${SHEET_NAME}" not found. Please check configuration.`);
        }

        sheet.clearContents();
        const headers = ["Select", "Date", "Start Time", "End Time", "Title", "Organizer", "Current Room", "Event ID"];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

        const now = new Date();
        const futureDate = new Date();
        futureDate.setDate(now.getDate() + DAYS_TO_CHECK);

        const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
        if (!calendar) {
            throw new Error(`Calendar with ID "${CALENDAR_ID}" not found.`);
        }

        const events = calendar.getEvents(now, futureDate);
        const data = [];

        events.forEach(event => {
            if (event.isAllDayEvent()) return;

            const startTime = event.getStartTime();
            const assignedRoom = getAssignedRoomName(event);

            data.push([
                false, // Checkbox for 'Select'
                Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd"),
                Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm"),
                Utilities.formatDate(event.getEndTime(), Session.getScriptTimeZone(), "HH:mm"),
                event.getTitle(),
                event.getCreators()[0] || "N/A",
                assignedRoom ? `Yes (${assignedRoom})` : "No",
                event.getId()
            ]);
        });

        if (data.length > 0) {
            sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
            sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).insertCheckboxes();
            sheet.hideColumns(headers.length); // Hide the "Event ID" column
            sheet.autoResizeColumns(1, headers.length - 1);
        }

        SpreadsheetApp.getUi().alert('Success', `Successfully populated ${data.length} meetings to the sheet.`);
        Logger.log("Meetings populated successfully.");

    } catch (e) {
        Logger.log(`Error populating meetings to sheet: ${e.message}`);
        SpreadsheetApp.getUi().alert('Error', `Failed to populate meetings: ${e.message}`);
    }
}

/**
 * Reads selected meetings from the Google Sheet and attempts to assign rooms.
 */
function assignRoomsFromSheet() {
    Logger.log("Assigning rooms from sheet...");
    try {
        const { sheet, values } = getSheetData();
        if (!values) return;

        let roomsAssignedCount = 0;
        const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

        values.forEach((row, index) => {
            const isSelected = row[0]; // Column A: "Select"
            const eventId = row[7];    // Column H: "Event ID"

            if (isSelected === true && eventId) {
                try {
                    const event = calendar.getEventById(eventId);
                    if (event) {
                        if (hasMeetingRoom(event)) {
                            Logger.log(`Event '${event.getTitle()}' already has a room. Skipping.`);
                            values[index][6] = `Already assigned (${getAssignedRoomName(event)})`;
                        } else {
                            const assigned = assignAvailableRoom(event);
                            if (assigned) {
                                values[index][6] = `Assigned: ${getAssignedRoomName(event)}`;
                                roomsAssignedCount++;
                            } else {
                                values[index][6] = "No room found";
                            }
                        }
                    } else {
                        values[index][6] = "Event not found";
                    }
                } catch (e) {
                    Logger.log(`Error processing event ID ${eventId}: ${e.message}`);
                    values[index][6] = `Error: ${e.message.substring(0, 50)}`;
                }
            }
        });

        // Batch update the "Current Room" column
        const statuses = values.map(row => [row[6]]);
        sheet.getRange(2, 7, statuses.length, 1).setValues(statuses);

        SpreadsheetApp.getUi().alert('Success', `Finished processing. Rooms assigned: ${roomsAssignedCount}`);
        Logger.log("Room assignment from sheet finished.");

    } catch (e) {
        Logger.log(`Error assigning rooms from sheet: ${e.message}`);
        SpreadsheetApp.getUi().alert('Error', `Failed to assign rooms from sheet: ${e.message}`);
    }
}

/**
 * Shifts the start and end times of selected meetings in the sheet by 5 minutes.
 */
function shiftSelectedMeetings() {
    Logger.log("Shifting selected meetings...");
    try {
        const { sheet, values } = getSheetData();
        if (!values) return;

        let shiftedCount = 0;
        const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

        values.forEach((row, index) => {
            const isSelected = row[0];
            const eventId = row[7];

            if (isSelected === true && eventId) {
                try {
                    const event = calendar.getEventById(eventId);
                    if (event) {
                        const startTime = event.getStartTime();
                        const minutes = startTime.getMinutes();

                        if (minutes === 0 || minutes === 30) {
                            const newStartTime = new Date(startTime.getTime() + 5 * 60 * 1000);
                            const newEndTime = new Date(event.getEndTime().getTime() + 5 * 60 * 1000);
                            event.setTime(newStartTime, newEndTime);

                            // Update the values array for batch update
                            values[index][2] = Utilities.formatDate(newStartTime, Session.getScriptTimeZone(), "HH:mm");
                            values[index][3] = Utilities.formatDate(newEndTime, Session.getScriptTimeZone(), "HH:mm");
                            shiftedCount++;
                            Logger.log(`Shifted event '${event.getTitle()}'`);
                        }
                    }
                } catch (e) {
                    Logger.log(`Error shifting event ID ${eventId}: ${e.message}`);
                }
            }
        });

        // Batch update the Start and End time columns in the sheet
        const times = values.map(row => [row[2], row[3]]);
        sheet.getRange(2, 3, times.length, 2).setValues(times);

        SpreadsheetApp.getUi().alert('Success', `Finished. Total meetings shifted: ${shiftedCount}.`);
        Logger.log("Finished shifting selected meetings.");
    } catch (e) {
        Logger.log(`Error shifting meetings: ${e.message}`);
        SpreadsheetApp.getUi().alert('Error', `Failed to shift meetings: ${e.message}`);
    }
}


// --- Automated Calendar Functions ---

/**
 * Main auto-assignment function. Assigns rooms to meetings that meet criteria.
 */
function assignMeetingRooms() {
    Logger.log("Starting automated meeting room assignment...");
    const now = new Date();
    const futureDate = new Date(now.getTime() + DAYS_TO_CHECK * 24 * 60 * 60 * 1000);
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
        Logger.log(`Calendar with ID "${CALENDAR_ID}" not found.`);
        return;
    }

    const events = calendar.getEvents(now, futureDate);
    Logger.log(`Found ${events.length} events to check in the next ${DAYS_TO_CHECK} days.`);
    const scriptOwnerEmail = Session.getActiveUser().getEmail().toLowerCase();
    let assignedCount = 0;

    events.forEach(event => {
        const eventTitle = event.getTitle();
        try {
            // Initial checks to skip events early
            if (event.isAllDayEvent()) {
                return; // Silently skip all-day events
            }
            if (event.getGuests().length < MINIMUM_PARTICIPANTS) {
                Logger.log(`Skipping '${eventTitle}': Fewer than ${MINIMUM_PARTICIPANTS} participants.`);
                return;
            }

            const eventOwnerEmail = (event.getCreators()[0] || "").toLowerCase();
            const isOwnedByMe = eventOwnerEmail === scriptOwnerEmail;
            let shouldProcess = false;

            // Logic for events owned by the script runner
            if (isOwnedByMe) {
                Logger.log(`Processing owned event: '${eventTitle}'`);
                if (!isInOfficeDay(event.getStartTime())) {
                    Logger.log(` -> Skipping '${eventTitle}': Not an in-office day.`);
                } else if (hasMeetingRoom(event)) {
                    Logger.log(` -> Skipping '${eventTitle}': Already has a room.`);
                } else {
                    shouldProcess = true;
                }
            } 
            // Logic for non-owned events if auto-assignment is enabled
            else if (AUTO_ASSIGN_NON_OWNED_MEETINGS) {
                Logger.log(`Processing non-owned event: '${eventTitle}'`);
                const participantCount = event.getGuests().length;
                if (participantCount >= NON_OWNED_MEETING_MAX_PARTICIPANTS) {
                    Logger.log(` -> Skipping '${eventTitle}': Exceeds max participants (${NON_OWNED_MEETING_MAX_PARTICIPANTS}) for non-owned meetings.`);
                } else if (!isInOfficeDay(event.getStartTime())) {
                    Logger.log(` -> Skipping '${eventTitle}': Not an in-office day.`);
                } else if (hasMeetingRoom(event)) {
                    Logger.log(` -> Skipping '${eventTitle}': Already has a room.`);
                } else {
                    shouldProcess = true;
                }
            }

            // If the event passes all checks, try to assign a room
            if (shouldProcess) {
                Logger.log(` -> Attempting to find a room for '${eventTitle}'...`);
                if (assignAvailableRoom(event)) {
                    assignedCount++;
                    Logger.log(` -> SUCCESS: Assigned a room to '${eventTitle}'`);
                } else {
                    Logger.log(` -> FAILED: Could not find an available room for '${eventTitle}'`);
                }
            }

        } catch (e) {
            Logger.log(`Error processing event '${eventTitle}': ${e.message}`);
        }
    });

    Logger.log(`Automated assignment finished. Rooms assigned: ${assignedCount}.`);
}


/**
 * Shifts owned calendar entries that start at HH:00 or HH:30 by 5 minutes.
 */
function shiftOwnedMeetingsBy5Mins() {
    Logger.log("Starting to shift owned meetings...");
    const now = new Date();
    const futureDate = new Date(now.getTime() + DAYS_TO_CHECK * 24 * 60 * 60 * 1000);
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    let shiftedCount = 0;

    const events = calendar.getEvents(now, futureDate);
    events.forEach(event => {
        try {
            if (event.isAllDayEvent()) return;

            const owner = (event.getCreators()[0] || "").toLowerCase();
            const isOwnedByUser = (owner === userEmail);
            const start = event.getStartTime();
            const startMinute = start.getMinutes();

            if (isOwnedByUser && (startMinute === 0 || startMinute === 30)) {
                const newStart = new Date(start.getTime() + 5 * 60000);
                const newEnd = new Date(event.getEndTime().getTime() + 5 * 60000);
                event.setTime(newStart, newEnd);
                shiftedCount++;
                Logger.log(`Shifted event: '${event.getTitle()}' to ${newStart.toLocaleTimeString()}`);
            }
        } catch (e) {
            Logger.log(`Error shifting event '${event.getTitle()}': ${e.message}`);
        }
    });
    SpreadsheetApp.getUi().alert('Success', `Finished. Total meetings shifted: ${shiftedCount}.`);
    Logger.log(`Finished shifting owned meetings. Total shifted: ${shiftedCount}.`);
}


// --- Helper Functions ---

/**
 * Gets the sheet and its data, handling common errors.
 * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, values: Object[][]}|{}}
 */
function getSheetData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
        SpreadsheetApp.getUi().alert('Error', `Sheet "${SHEET_NAME}" not found.`);
        return {};
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        SpreadsheetApp.getUi().alert('Info', 'No meetings found in the sheet to process.');
        return {};
    }
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    return { sheet, values: dataRange.getValues() };
}

/**
 * Checks if the given date is an "in-office" day.
 * @param {Date} date The date to check.
 * @returns {boolean} True if it's an in-office day.
 */
function isInOfficeDay(date) {
    return IN_OFFICE_DAYS.includes(date.getDay());
}

/**
 * Checks if an event already has a meeting room assigned from the AVAILABLE_ROOMS list.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The calendar event.
 * @returns {boolean} True if a meeting room is assigned.
 */
function hasMeetingRoom(event) {
    return event.getGuestList().some(guest => AVAILABLE_ROOMS.hasOwnProperty(guest.getEmail().toLowerCase()));
}

/**
 * Finds and assigns an available meeting room to an event.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The calendar event to update.
 * @returns {boolean} True if a room was assigned.
 */
function assignAvailableRoom(event) {
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const eventTitle = event.getTitle();

    for (const roomEmail of Object.keys(AVAILABLE_ROOMS)) {
        const roomName = AVAILABLE_ROOMS[roomEmail];
        try {
            const roomCalendar = CalendarApp.getCalendarById(roomEmail);
            if (roomCalendar) {
                const existingEvents = roomCalendar.getEvents(startTime, endTime);
                if (existingEvents.length === 0) {
                    event.addGuest(roomEmail);
                    return true;
                } else {
                    Logger.log(`   - Room '${roomName}' is occupied for '${eventTitle}'.`);
                }
            } else {
                Logger.log(`   - Could not access calendar for room: ${roomName}`);
            }
        } catch (e) {
            Logger.log(`   - Error processing room ${roomName}: ${e.message}`);
        }
    }
    return false;
}

/**
 * Gets the friendly name of an assigned room from an event.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The event to check.
 * @returns {string|null} The name of the assigned room, or null if none.
 */
function getAssignedRoomName(event) {
    const guests = event.getGuestList();
    for (const guest of guests) {
        const email = guest.getEmail().toLowerCase();
        if (AVAILABLE_ROOMS.hasOwnProperty(email)) {
            return AVAILABLE_ROOMS[email];
        }
    }
    return null;
}


// --- Trigger Setup ---

/**
 * Sets up a daily time-driven trigger to run the 'assignMeetingRooms' function.
 * Run this once manually to establish the trigger.
 */
function setupTrigger() {
    // Delete existing triggers to avoid duplicates
    ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getHandlerFunction() === 'assignMeetingRooms') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // Create a new trigger to run daily between 1 AM and 2 AM
    ScriptApp.newTrigger('assignMeetingRooms')
        .timeBased()
        .everyDays(1)
        .atHour(1)
        .create();

    Logger.log("Daily trigger for 'assignMeetingRooms' has been set up.");
    SpreadsheetApp.getUi().alert('Trigger Setup', 'Daily trigger for automated room assignment has been set up (runs between 1-2 AM).');
}