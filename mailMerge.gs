
/** 
 * This script enhances the original script of Martin Hawksey by:
 * Adding support for CC
 * ability to send the emails a few minutes/hours later.
 * This creates a trigger to execute the script to send email at end of time period. 
 * The recipent info is stored in script properties of the script for trigger to use. 
 * Be aware of the properties quota for any limits https://developers.google.com/apps-script/guides/services/quotas

 * To learn how to use this script, refer to the documentation:
 * https://developers.google.com/apps-script/samples/automations/mail-merge


*/

/**
 * @OnlyCurrentDoc
 */

/**
 * Change these to match the column names you are using for email
 * recipient addresses, CC addresses, and the email sent column.
 */
const RECIPIENT_COL  = "RECIPIENT";
const CC_COL         = "CC"; // New column for CC recipients
const EMAIL_SENT_COL = "Sent";
const EMAIL_DELAY = 5 * 60 * 1000 // 5 Minutes (5*60*1000 ms)

// Name of the function that will be triggered to send emails.
const TRIGGER_FUNCTION_NAME = 'sendTriggeredEmails';

/**
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Emails Immediately', 'sendEmails')
      .addItem('Send Emails with 5-Min Delay', 'scheduleEmails')
      .addSeparator()
      .addItem('Cancel Scheduled Send', 'cancelScheduledSend')
      .addToUi();
}

/**
 * Deletes all existing project triggers that are set to run the sending function.
 * This prevents multiple triggers from being set accidentally.
 */
function deleteExistingTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Schedules emails to be sent after a 5 minute delay.
 * It prepares the emails, stores their data, and creates a time-based trigger.
 */
function scheduleEmails() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const subjectLine = Browser.inputBox("Schedule Mail Merge",
                                     "Enter the subject line of the Gmail draft to use as a template.\n\n" +
                                     "Emails will be sent in 5 minutes.",
                                     Browser.Buttons.OK_CANCEL);
  if (subjectLine === "cancel" || subjectLine === "") {
    return;
  }

  // Clear any previously scheduled sends to start fresh
  deleteExistingTriggers_();
  PropertiesService.getScriptProperties().deleteProperty('emails_to_send');

  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const emailsToSchedule = [];
  const statuses = [];

  obj.forEach(function(row, rowIdx){
    // Only schedule emails if the status cell is blank and a recipient is listed
    if (row[EMAIL_SENT_COL] === '' && row[RECIPIENT_COL]) {
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        const mailOptions = {
          htmlBody: msgObj.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        };
        // Add CC if the column exists and has a value
        if (row[CC_COL]) {
          mailOptions.cc = row[CC_COL];
        }

        emailsToSchedule.push({
          rowIndex: rowIdx + 2, // +2 for 1-based index and header row
          recipient: row[RECIPIENT_COL],
          subject: msgObj.subject,
          text: msgObj.text,
          options: mailOptions
        });
        statuses.push(['Scheduled']);
      } catch(e) {
        statuses.push([e.message]);
      }
    } else {
      statuses.push([row[EMAIL_SENT_COL]]);
    }
  });

  if (emailsToSchedule.length === 0) {
    Browser.msgBox('No emails to schedule. Check that recipient emails are listed and the "Email Sent" column is empty for those rows.');
    return;
  }

  // Store the email data for the trigger function
  PropertiesService.getScriptProperties().setProperty('emails_to_send', JSON.stringify(emailsToSchedule));

  // Create a trigger to run in 5 minutes
  ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .timeBased()
      .at(new Date(new Date().getTime() + EMAIL_DELAY))
      .create();

  // Update the sheet with "Scheduled" status
  sheet.getRange(2, emailSentColIdx + 1, statuses.length).setValues(statuses);

  SpreadsheetApp.getUi().alert(`Scheduled ${emailsToSchedule.length} email(s) to be sent in 5 minutes.\n\nTo cancel, use "Mail Merge > Cancel Scheduled Send".`);
}

/**
 * This function is executed by a time-based trigger to send the scheduled emails.
 */
function sendTriggeredEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = PropertiesService.getScriptProperties().getProperty('emails_to_send');
  if (!data) {
    // No data to process, so clean up triggers and exit.
    deleteExistingTriggers_();
    return;
  }

  const emails = JSON.parse(data);
  const heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  emails.forEach(email => {
    let status;
    try {
      GmailApp.sendEmail(email.recipient, email.subject, email.text, email.options);
      status = [new Date()];
    } catch (e) {
      status = [e.message];
    }
    // Update status for the corresponding row
    sheet.getRange(email.rowIndex, emailSentColIdx + 1).setValues([status]);
  });

  // Clean up after sending
  PropertiesService.getScriptProperties().deleteProperty('emails_to_send');
  deleteExistingTriggers_();
}

/**
 * Cancels any pending scheduled email sends.
 */
function cancelScheduledSend() {
  deleteExistingTriggers_();
  PropertiesService.getScriptProperties().deleteProperty('emails_to_send');

  // Clear "Scheduled" status from the sheet
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  if (emailSentColIdx === -1) {
    SpreadsheetApp.getUi().alert('Cannot find "Email Sent" column to clear statuses.');
    return;
  }

  const statusesToClear = data.map(row => {
    return (row[emailSentColIdx] === 'Scheduled') ? [''] : [row[emailSentColIdx]];
  });

  sheet.getRange(2, emailSentColIdx + 1, statusesToClear.length).setValues(statusesToClear);

  SpreadsheetApp.getUi().alert('Scheduled email send has been canceled.');
}

/**
 * Sends emails from sheet data immediately.
 * This is the original function, modified to include the CC feature.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 */
function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  if (!subjectLine) {
    subjectLine = Browser.inputBox("Mail Merge",
                                   "Type or copy/paste the subject line of the Gmail " +
                                   "draft message you would like to mail merge with:",
                                   Browser.Buttons.OK_CANCEL);
    if (subjectLine === "cancel" || subjectLine == "") {
      return;
    }
  }

  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  const out = [];

  obj.forEach(function(row, rowIdx) {
    if (row[EMAIL_SENT_COL] == '' && row[RECIPIENT_COL]) {
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        const mailOptions = {
          htmlBody: msgObj.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        };
        // Add CC if the column exists and has a value
        if (row[CC_COL]) {
          mailOptions.cc = row[CC_COL];
        }
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, mailOptions);
        out.push([new Date()]);
      } catch (e) {
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  if (out.length > 0) {
    sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
  }
}

/**
 * Helper functions from the original script below. No changes were needed here.
 */

function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    const msg = draft.getMessage();
    const allInlineImages = draft.getMessage().getAttachments({
      includeInlineImages: true,
      includeAttachments: false
    });
    const attachments = draft.getMessage().getAttachments({
      includeInlineImages: false
    });
    const htmlBody = msg.getBody();
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];
    const inlineImagesObj = {};
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);
    return {
      message: {
        subject: subject_line,
        text: msg.getPlainBody(),
        html: htmlBody
      },
      attachments: attachments,
      inlineImages: inlineImagesObj
    };
  } catch (e) {
    throw new Error("Oops - can't find Gmail draft with subject: '" + subject_line + "'");
  }

  function subjectFilter_(subject_line) {
    return function(element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

function fillInTemplateFromObject_(template, data) {
  let template_string = JSON.stringify(template);
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return JSON.parse(template_string);
}

function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
};