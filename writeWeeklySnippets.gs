/**
 * @OnlyCurrentDoc
 * This script adds a custom menu to a Google Doc to generate weekly work snippets. It fetches
 * All your calendar entries, 
 * Sent Emails
 * Tasks completed in the last week,
 * Edited Google Drive files,
 * Jira tasks created or completed in the last week
 * 
 * Note: This script requires the Google Calendar, Gmail, Google Drive, and Tasks APIs to be enabled in your Google Cloud project.
 * This script also requires the Jira API to be enabled and configured with your Jira domain, user email, and API token.
 * It is important to replace the placeholder values for JIRA_DOMAIN, JIRA_USER_EMAIL, and JIRA_API_TOKEN with your actual Jira configuration.
 * The script cannot work inside a table in the Google Doc, so it will prompt the user to place the cursor outside of any table before generating snippets.
 */

// --- CONFIGURATION ---
const DAYS_TO_LOOK_BACK = 7;
const DAYS_TO_LOOK_FORWARD = 7;

// --- JIRA CONFIGURATION ---
// IMPORTANT: You must fill these out for Jira integration to work.
const JIRA_DOMAIN = "YOUR_DOMAIN.atlassian.net"; // Replace with your Jira domain (e.g., "your-company.atlassian.net")
const JIRA_USER_EMAIL = "your-email@example.com"; // Replace with your email used for Jira
const JIRA_API_TOKEN = "YOUR_JIRA_API_TOKEN"; // Replace with your Jira API token from https://id.atlassian.com/manage-profile/security/api-tokens.

/**
 * Creates a custom menu in the Google Doc UI when the document is opened.
 */
function onOpen() {
    DocumentApp.getUi()
        .createMenu('Snippets')
        .addItem('Generate Weekly Snippets', 'generateSnippets')
        .addToUi();
}


/**
 * Checks if the cursor is currently inside a table element.
 * @param {GoogleAppsScript.Document.Position} cursor The current cursor position.
 * @returns {boolean} True if the cursor is in a table, false otherwise.
 */
function isCursorInTable(cursor) {
    if (!cursor) {
        return false;
    }
    let parent = cursor.getElement().getParent();
    while (parent) {
        const type = parent.getType();
        if (type === DocumentApp.ElementType.TABLE_CELL) {
            return true;
        }
        if (type === DocumentApp.ElementType.BODY_SECTION) {
            return false;
        }
        parent = parent.getParent();
    }
    return false;
}

/**
 * Main function to generate and insert the snippets into the document.
 */
function generateSnippets() {
    const ui = DocumentApp.getUi();
    const cursor = DocumentApp.getActiveDocument().getCursor();
    if (!cursor) {
        ui.alert('Could not find a cursor position. Please click somewhere in the document first.');
        return;
    }

    if (isCursorInTable(cursor)) {
        ui.alert('Please place your cursor outside of a table to generate snippets.');
        return;
    }

    // Check if Jira config is placeholder
    if (JIRA_DOMAIN === "YOUR_DOMAIN.atlassian.net" || JIRA_USER_EMAIL === "your-email@example.com" || JIRA_API_TOKEN === "YOUR_JIRA_API_TOKEN") {
        ui.alert("Jira integration is not configured. Please update the JIRA_DOMAIN, JIRA_USER_EMAIL, and JIRA_API_TOKEN constants in the script.");
    }


    const lastWeekSnippets = getLastWeekSnippets();
    const nextWeekSnippets = getNextWeekSnippets();

    insertSnippets(cursor, "Last Week's Snippets", lastWeekSnippets);
    insertSnippets(cursor, "Next Week's Plan", nextWeekSnippets);
}

/**
 * Fetches data from various services for the last week.
 * @returns {Array<string|Object>} An array of snippet strings or objects with text and URL.
 */
function getLastWeekSnippets() {
    const today = new Date();
    const pastDate = new Date();
    pastDate.setDate(today.getDate() - DAYS_TO_LOOK_BACK);

    const pastDateString = Utilities.formatDate(pastDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

    // Calendar Events (skipping all-day events)
    const calendarEvents = CalendarApp.getDefaultCalendar().getEvents(pastDate, today);
    const eventTitles = calendarEvents
        .filter(event => !event.isAllDayEvent())
        .map(event => `- (Event) ${event.getTitle()}`);

    // Completed Tasks
    let completedTasks = [];
    try {
        const taskLists = Tasks.Tasklists.list().items;
        if (taskLists && taskLists.length > 0) {
            taskLists.forEach(taskList => {
                const tasks = Tasks.Tasks.list(taskList.id, {
                    showCompleted: true,
                    completedMin: pastDate.toISOString(),
                    completedMax: today.toISOString()
                }).items;
                if (tasks) {
                    completedTasks = completedTasks.concat(tasks.map(task => `- (Task) ${task.title}`));
                }
            });
        }
    } catch (e) {
        console.error("Error fetching Google Tasks: " + e.toString());
    }

    // Sent Emails
    const sentEmails = GmailApp.search(`is:sent after:${pastDateString}`).slice(0, 20); // Limit to 20 recent emails
    const emailSubjects = sentEmails.map(thread => `- (Email) Sent: ${thread.getFirstMessageSubject()}`);

    // Edited Google Drive Files (with hyperlinks)
    const driveQuery = `'me' in writers and modifiedDate > '${pastDate.toISOString()}'`;
    const editedFiles = DriveApp.searchFiles(driveQuery);
    let editedFileNames = [];
    while (editedFiles.hasNext()) {
        const file = editedFiles.next();
        editedFileNames.push({ text: `- (Doc) Edited: ${file.getName()}`, url: file.getUrl() });
    }

    // Jira Tasks
    const createdJiraTasks = getCreatedJiraTasks();
    const completedJiraTasks = getCompletedJiraTasks();

    return [].concat(eventTitles, completedTasks, emailSubjects, editedFileNames, createdJiraTasks, completedJiraTasks);
}

/**
 * Fetches upcoming tasks and events for the next week.
 * @returns {Array<string|Object>} An array of snippet strings or objects.
 */
function getNextWeekSnippets() {
    const today = new Date();
    const futureDate = new Date();
    futureDate.setDate(today.getDate() + DAYS_TO_LOOK_FORWARD);

    // Calendar Events (skipping all-day events)
    const calendarEvents = CalendarApp.getDefaultCalendar().getEvents(today, futureDate);
    const eventTitles = calendarEvents
        .filter(event => !event.isAllDayEvent())
        .map(event => `- (Event) ${event.getTitle()}`);

    let upcomingTasks = [];
    try {
        const taskLists = Tasks.Tasklists.list().items;
        if (taskLists && taskLists.length > 0) {
            taskLists.forEach(taskList => {
                const tasks = Tasks.Tasks.list(taskList.id, {
                    showCompleted: false,
                    dueMax: futureDate.toISOString()
                }).items;
                if (tasks) {
                    const futureTasks = tasks.filter(task => task.due && new Date(task.due) > today);
                    upcomingTasks = upcomingTasks.concat(futureTasks.map(task => `- (Task) ${task.title}`));
                }
            });
        }
    } catch (e) {
        console.error("Error fetching upcoming Google Tasks: " + e.toString());
    }

    return eventTitles.concat(upcomingTasks);
}

/**
 * Fetches tasks created by the user in the last week from Jira.
 * @returns {string[]} An array of Jira snippet strings.
 */
function getCreatedJiraTasks() {
    if (JIRA_DOMAIN === "YOUR_DOMAIN.atlassian.net") {
        return ["- (Jira) Not configured."];
    }

    const jql = `reporter = currentUser() AND created >= -${DAYS_TO_LOOK_BACK}d`;
    const url = `https://${JIRA_DOMAIN}/rest/api/2/search`;
    const encodedToken = Utilities.base64Encode(`${JIRA_USER_EMAIL}:${JIRA_API_TOKEN}`);

    const params = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
            'Authorization': `Basic ${encodedToken}`
        },
        'payload': JSON.stringify({
            'jql': jql,
            'fields': ['summary']
        }),
        'muteHttpExceptions': true
    };

    try {
        const response = UrlFetchApp.fetch(url, params);
        const json = JSON.parse(response.getContentText());
        if (json.issues) {
            return json.issues.map(issue => `- (Jira) Created: ${issue.key} - ${issue.fields.summary}`);
        } else {
            console.error("Jira API Error (Created): " + json.errorMessages.join(" "));
            return [`- (Jira) Error on Created Tasks: ${json.errorMessages.join(" ")}`];
        }
    } catch (e) {
        console.error("Failed to fetch created tasks from Jira: " + e.toString());
        return ["- (Jira) Failed to connect for created tasks."];
    }
}

/**
 * Fetches tasks completed by the user in the last week from Jira.
 * @returns {string[]} An array of Jira snippet strings.
 */
function getCompletedJiraTasks() {
    if (JIRA_DOMAIN === "YOUR_DOMAIN.atlassian.net") {
        return []; // Return empty if not configured
    }

    const jql = `assignee = currentUser() AND status = Done AND resolved >= -${DAYS_TO_LOOK_BACK}d`;
    const url = `https://${JIRA_DOMAIN}/rest/api/2/search`;
    const encodedToken = Utilities.base64Encode(`${JIRA_USER_EMAIL}:${JIRA_API_TOKEN}`);

    const params = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
            'Authorization': `Basic ${encodedToken}`
        },
        'payload': JSON.stringify({
            'jql': jql,
            'fields': ['summary']
        }),
        'muteHttpExceptions': true
    };

    try {
        const response = UrlFetchApp.fetch(url, params);
        const json = JSON.parse(response.getContentText());
        if (json.issues) {
            return json.issues.map(issue => `- (Jira) Completed: ${issue.key} - ${issue.fields.summary}`);
        } else {
            console.error("Jira API Error (Completed): " + json.errorMessages.join(" "));
            return [`- (Jira) Error on Completed Tasks: ${json.errorMessages.join(" ")}`];
        }
    } catch (e) {
        console.error("Failed to fetch completed tasks from Jira: " + e.toString());
        return ["- (Jira) Failed to connect for completed tasks."];
    }
}


/**
 * Inserts a title and a list of items at the cursor's position.
 * @param {GoogleAppsScript.Document.Position} cursor The position to insert the text.
 * @param {Array<string|Object>} items The list of items to insert. Can be strings or objects with {text, url}.
 */
function insertSnippets(cursor, title, items) {
    const body = DocumentApp.getActiveDocument().getBody();
    const cursorElement = cursor.getElement();
    const index = body.getChildIndex(cursorElement);

    const titlePara = body.insertParagraph(index + 1, title);
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    let insertionIndex = index + 2;

    if (items && items.length > 0) {
        items.forEach(item => {
            if (item) { // Ensure item is not null or undefined
                let text, url;
                if (typeof item === 'string') {
                    text = item;
                } else {
                    text = item.text;
                    url = item.url;
                }

                const listItem = body.insertListItem(insertionIndex, text);
                listItem.setGlyphType(DocumentApp.GlyphType.BULLET);

                if (url) {
                    // Set link for the entire list item. This is the corrected line.
                    listItem.setLinkUrl(url);
                }

                insertionIndex++;
            }
        });
        body.insertParagraph(insertionIndex, '');
    } else {
        body.insertParagraph(insertionIndex, 'No items found for this period.');
        body.insertParagraph(insertionIndex + 1, '');
    }
}