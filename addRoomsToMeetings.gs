

/**
 * Google Apps Script to check upcoming calendar meetings and assign rooms.
 *
 * This script looks for meetings:
 * - Owned by the user running the script.
 * - On days when the user is marked as "in office" (defined in inOfficeDays function).
 * - That do not currently have a meeting room assigned.
 *
 * It then attempts to find and assign an available meeting room within the specified office location and floor.
 */

// --- Configuration ---
const CALENDAR_ID = "primary";                          // Use "primary" for your main calendar
const DAYS_TO_CHECK = 30;                               // Check for meetings in the next 30 days

// Define your available meeting rooms. These are typically email addresses like "room-name@yourdomain.com".
// You'll need to get the exact resource IDs for your rooms.
const AVAILABLE_ROOMS = [

    // Add more room resource IDs as needed
];

// --- Main Function ---
function assignMeetingRooms() {
    Logger.log("Starting meeting room assignment script...");

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

    Logger.log("Meeting room assignment script finished.");
}

// --- Helper Functions ---

/**
 * Processes a single calendar event to check if it needs a room and assign one.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The calendar event to process.
 */
function processEvent(event) {
    const eventTitle = event.getTitle();
    const eventOwnerEmail = event.getCreators().length > 0 ? event.getCreators()[0] : null;
    const eventStartTime = event.getStartTime();
    const eventEndTime = event.getEndTime();

    // 1. Check if owned by me (the script runner)
    // Note: getCreators() returns an array of emails. If the event is created by a user
    // whose primary email matches the script owner, it's considered owned.
    // This might need refinement based on your organization's event ownership rules.
    const scriptOwnerEmail = Session.getActiveUser().getEmail();
    const isOwnedByMe = eventOwnerEmail && eventOwnerEmail.toLowerCase() === scriptOwnerEmail.toLowerCase();

    if (!isOwnedByMe) {
        Logger.log(`Skipping event '${eventTitle}': Not owned by me.`);
        return;
    }

    if (event.isAllDayEvent()) {
        Logger.log("Event: " + event.getTitle() + " is all day event:")
        return;
    }

    // 2. Check if on a day I am in office
    if (!isInOfficeDay(eventStartTime)) {
        Logger.log(`Skipping event '${eventTitle}': Not an in-office day.`);
        return;
    }

    // 3. Check if it already has a meeting room assigned
    if (hasMeetingRoom(event)) {
        //Logger.log(`Skipping event '${eventTitle}': Already has a meeting room.`);
        return;
    }

    Logger.log(`Processing eligible event: '${eventTitle}' on ${eventStartTime.toLocaleString()}`);

    // Attempt to find and assign a room
    const assigned = assignAvailableRoom(event, eventStartTime, eventEndTime);
    if (assigned) {
        Logger.log(`Successfully assigned a room to: '${eventTitle}'`);
    } else {
        Logger.log(`Could not find an available room for: '${eventTitle}'`);
    }
}

/**
 * Checks if the given date is an "in-office" day based on a specific list of days.
 * @param {Date} date The date to check.
 * @returns {boolean} True if the date is an in-office day, false otherwise.
 */
function isInOfficeDay(date) {
    const dayOfWeek = date.getDay(); // Sunday - 0, Monday - 1, ..., Saturday - 6
    // Define your in-office days here.  Modify this array to match your schedule.
    const inOfficeDays = [1, 3, 4]; // Monday (1), Wednesday (3), Thursday (4)

    return inOfficeDays.includes(dayOfWeek);
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
    //const guestEmails = event.getGuestList();
    const guestEmails = event.getGuests();
    Logger.log("Event Guests: " + event.getGuests());
    for (let i = 0; i < guestEmails.length; i++) {
        const guestEmail = guestEmails[i].toLowerCase(); // Convert to lowercase for case-insensitive comparison
        if (AVAILABLE_ROOMS.includes(guestEmail)) {
            Logger.log(`Event '${event.getTitle()}' already has room: ${guestEmail}`);
            return true;
        }
    }
    return false;
}


/**
 * Finds and assigns an available meeting room to an event.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The calendar event to update.
 * @param {Date} startTime The start time of the event. // Not directly used in the final version of this function, but kept for context if needed.
 * @param {Date} endTime The end time of the event.   // Not directly used in the final version of this function, but kept for context if needed.
 * @returns {boolean} True if a room was assigned, false otherwise.
 */
/**function assignAvailableRoom(event, startTime, endTime) { // startTime and endTime are not explicitly used here as addGuest handles time internally
  for (const roomEmail of AVAILABLE_ROOMS) {
    try {
      // Adding a guest with a resource email attempts to book the resource.
      // Google Calendar handles the availability check internally when addGuest is called.
      event.addGuest(roomEmail);
      
      Logger.log(`Attempted to add room ${roomEmail} to event '${event.getTitle()}'`);
      // If addGuest succeeds without throwing an error, we assume the room was assigned (or an attempt was made).
      return true; // Room assigned (or attempt made to assign)
    } catch (e) {
      // This catch block handles cases where:
      // 1. The room is unavailable for the event time.
      // 2. The roomEmail is not a valid resource email in your domain.
      // 3. There are other permission issues.
      Logger.log(`Could not add room ${roomEmail} to event '${event.getTitle()}': ${e.message}`);
      continue; // Try the next room in the list
    }
  }
  return false; // No room could be assigned after trying all available rooms
}
*/



/**
 * Finds and assigns an available meeting room to an event.
 * Checks room availability by getting the room's calendar and querying its events.
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event The calendar event to update.
 * @param {Date} startTime The start time of the event.
 * @param {Date} endTime The end time of the event.
 * @returns {boolean} True if an available room was assigned, false otherwise.
 */
function assignAvailableRoom(event, startTime, endTime) {
    for (const roomEmail of AVAILABLE_ROOMS) {
        try {
            // 1. Get the Calendar object for the room using its email as the ID.
            const roomCalendar = CalendarApp.getCalendarById(roomEmail);

            // Ensure the room email corresponds to a valid calendar that we can access.
            if (!roomCalendar) {
                Logger.log(`Could not retrieve calendar for room email: ${roomEmail}. It might not be a valid resource email or permissions are missing.`);
                continue; // Try the next room
            }

            // 2. Check if the room calendar has any events overlapping with the meeting time.
            const existingEventsForRoom = roomCalendar.getEvents(startTime, endTime);

            if (existingEventsForRoom.length === 0) {
                // If no existing events are found, the room is available.
                // 3. Attempt to add the room as a guest to the event.
                event.addGuest(roomEmail);
                Logger.log(`Successfully assigned available room ${roomEmail} to event '${event.getTitle()}'`);
                return true; // Room found and assigned
            } else {
                // Room is occupied
                Logger.log(`Room ${roomEmail} is occupied for event '${event.getTitle()}' from ${startTime.toLocaleString()} to ${endTime.toLocaleString()}.`);
                // Continue to the next room
            }
        } catch (e) {
            // This catch block handles potential errors during the process, such as:
            // - Problems getting the calendar by ID (though handled by `!roomCalendar` check too).
            // - Errors when adding the guest (e.g., transient network issues, permissions on the event itself).
            Logger.log(`Error when processing room ${roomEmail} for event '${event.getTitle()}': ${e.message}`);
            continue; // Try the next room
        }
    }
    return false; // No available room could be assigned after trying all
}


// --- Setup Trigger Function (Run this once to set up the daily trigger) ---
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
}