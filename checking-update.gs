const EMAIL_LABEL = "bank-update";
const SHEET_LABEL = "emails-test";

/**
 * Saves the target email label to the bound Spreadsheet.
 * @returns {void}
 */
function saveLabelledEmailsToSheets() {
  const messages = getLabelledEmails();
  if (messages.length <= 0) return; // return if there's no emails in that label

  // get the target sheet
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LABEL);
  const data_height = sheet.getDataRange().getHeight();
  // get ids for the messages to filter new ones
  const emailID = sheet.getRange(2, 2, data_height, 1).getValues().flat();

  const newData = [];
  messages.forEach((m) => {
    // add new emails to the array
    if (!emailID.includes(m.getId())) {
      newData.push([
        m.getDate(),
        m.getId(),
        m.getFrom(),
        m.getSubject(),
        m.getPlainBody(),
      ]);
    }
    // mark as read so we don't load it on the next run
    m.markRead();
  });

  // if new data is empty return
  if (newData.length <= 0) return;

  //
  try {
    sheet
      .getRange(data_height + 1, 1, newData.length, newData[0].length)
      .setValues(newData);
  } catch (e) {
    // mark emails as unread if we failed to add to the sheet for some reason
    messages.forEach((m) => m.markUnread());
    throw e;
  }
}

/**
 *  * Returns emails from the Gmail inbox in the given Gmail Label
 * @returns {Array<GmailMessage>}
 */
function getLabelledEmails() {
  // Fetch email threads from the Gmail inbox with a label search query
  const labelThreads = GmailApp.search(`is:unread label:${EMAIL_LABEL}`);
  // get all the messages for the current batch of threads
  const messages = GmailApp.getMessagesForThreads(labelThreads);
  // Flatten to a single array (get messages returns an array of arrays)
  return messages.flat();
}
