const BAL_EMAIL_LABEL = "balance-log";
const BAL_SHEET_LABEL = "emails-test";

/**
 * Saves the target email label to the bound Spreadsheet.
 * @returns {void}
 */
function saveLabelledEmailsToSheets() {
  // Part 1: Balance Log
  try {
    processBalanceLogEmails();
  } catch (e) {
    // Log the error and continue (we don't want to stop the other processes)
    Logger.log(e);
  }

  // Part 2: Transaction Log
}

/**
 * Processes balance log emails and updates the balance log sheet.
 * @returns {void}
 */
function processBalanceLogEmails() {
  const messages = getLabelledEmails(`is:unread label:${BAL_EMAIL_LABEL}`);
  if (messages.length <= 0) return; // return if there's no emails in that label

  // get the target sheet
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BAL_SHEET_LABEL);
  const data_height = sheet.getDataRange().getHeight();
  // get ids for the messages to filter new ones
  const emailID = sheet.getRange(2, 1, data_height, 1).getValues().flat();

  // regex to match the date and balance
  const regex = /as of (\d{1,2}\/\d{1,2}\/\d{4} \d{1,2}:\d{2}:\d{2} [AP]M) is \$([0-9,.]+),/;

  const newData = [];
  messages.forEach((m) => {
    // add new emails to the array
    if (!emailID.includes(m.getId())) {
      const message_body = m.getPlainBody();

      // match the date and balance from the email body
      const match = regex.exec(message_body.replaceAll("\r","").replaceAll("\n"," "));
      const message_date = match[1];
      const message_balance = match[2].replaceAll(",",""); 

      newData.push([
        m.getId(),
        m.getDate(),
        message_date,
        m.getSubject(),
        message_balance,
      ]);
    }
    // read emails so we don't load it on the next run
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
    // unread emails if we failed to add to the sheet for some reason
    messages.forEach((m) => m.markUnread());
    throw e;
  }
}

/**
 *  * Returns emails from the Gmail inbox in the given Gmail Label
 * @param {string} query - Gmail search query
 * @returns {Array<GmailMessage>}
 */
function getLabelledEmails(query = "") {
  // Fetch email threads from the Gmail inbox with a label search query
  const labelThreads = GmailApp.search(query);
  // get all the messages for the current batch of threads
  const messages = GmailApp.getMessagesForThreads(labelThreads);
  // Flatten to a single array (get messages returns an array of arrays)
  return messages.flat();
}
