const BAL_EMAIL_LABEL = "ssb-balance-log";
const BAL_SHEET_LABEL = "checking-balance-log";

const TRN_EMAIL_LABEL = "ssb-transaction-log";
const TRN_SHEET_LABEL = "checking-transaction-log";

/**
 * Saves the target email label to the bound Spreadsheet.
 * @returns {void}
 */
function saveLabelledEmailsToSheets() {
  // Part 1: Balance Log
  // Get all unread emails with the label $BAL_EMAIL_LABEL
  const bal_emails = getLabelledEmails(`is:unread label:${BAL_EMAIL_LABEL}`);
  try {
    processBalanceLogEmails(bal_emails);
  } catch (e) {
    // unread emails if we failed to add to the sheet for some reason
    bal_emails.forEach((m) => m.markUnread());
    // Log the error and continue (we don't want to stop the other processes)
    console.log(e);
  }

  // Part 2: Transaction Log
  // Get all unread emails with the label $TRN_EMAIL_LABEL
  const trn_emails = getLabelledEmails(`is:unread label:${TRN_EMAIL_LABEL}`);
  try {
    processTransactionLogEmails(trn_emails);
  } catch (e) {
    // unread emails if we failed to add to the sheet for some reason
    trn_emails.forEach((m) => m.markUnread());
    // Log the error and continue (we don't want to stop the other processes)
    console.log(e);
  }
}

/**
 * Processes balance log emails and updates the balance log sheet.'
 * @param {Array<GmailMessage>} messages - Array of Gmail messages
 * @returns {void}
 */
function processBalanceLogEmails(messages = []) {
  if (messages.length <= 0) return; // return if there's no emails in that label

  // get the target sheet
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BAL_SHEET_LABEL);
  const data_height = sheet.getDataRange().getHeight();
  // get ids for the messages to filter new ones
  const emailID = sheet.getRange(2, 1, data_height, 1).getValues().flat();

  // regex to match the date and balance
  const regex =
    /(\d{4}) as of (\d{1,2}\/\d{1,2}\/\d{4} \d{1,2}:\d{2}:\d{2} [AP]M) is \$([0-9,.]+),/;

  const newData = [];
  messages.forEach((m) => {
    // add new emails to the array
    if (!emailID.includes(m.getId())) {
      const message_body = m.getPlainBody();

      // match the date and balance from the email body
      const match = regex.exec(
        message_body.replaceAll("\r", "").replaceAll("\n", " ")
      );
      const message_account = match[1];
      const message_date = match[2];
      const message_balance = match[3].replaceAll(",", "");

      newData.push([
        m.getId(),
        m.getDate(),
        message_date,
        message_account,
        m.getSubject(),
        message_balance,
      ]);
    }
    // read emails so we don't load it on the next run
    m.markRead();
  });

  // if new data is empty return
  if (newData.length <= 0) return;

  sheet
    .getRange(data_height + 1, 1, newData.length, newData[0].length)
    .setValues(newData);
}

/**
 * Processes transaction log emails and updates the balance log sheet.'
 * @param {Array<GmailMessage>} messages - Array of Gmail messages
 * @returns {void}
 */
function processTransactionLogEmails(messages = []) {
  if (messages.length <= 0) return; // return if there's no emails in that label

  // get the target sheet
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRN_SHEET_LABEL);
  const data_height = sheet.getDataRange().getHeight();
  // get ids for the messages to filter new ones
  const emailID = sheet.getRange(2, 7, data_height, 1).getValues().flat();

  // regex to match the date and balance
  const regex =
    / A (.+?) (credit|debit) of \$([0-9,.]+) was \w+ \w+ your account \*(\d{4}) on (\d{1,2}\/\d{1,2}\/\d{4} \d{1,2}:\d{2}:\d{2} [AP]M),/;

  const newData = [];
  messages.forEach((m) => {
    // add new emails to the array
    if (!emailID.includes(m.getId())) {
      const message_body = m.getPlainBody();

      // match the date and balance from the email body
      const match = regex.exec(
        message_body.replaceAll("\r", "").replaceAll("\n", " ")
      );
      const description = match[1];
      const type = match[2];
      const amount = match[3].replaceAll(",", "");
      const post_account = match[4];
      const post_date = match[5];

      // define a variable sign which is 1 for credit and -1 for debit
      const sign = type === "credit" ? 1 : -1;

      newData.push([
        "055001096",
        post_account,
        "Checking",
        "BUSINESS INTEREST CHECKING",
        post_date,
        "",
        m.getId(),
        sign * amount,
        description,
        type,
        "",
      ]);
    }
    // read emails so we don't load it on the next run
    m.markRead();
  });

  // if new data is empty return
  if (newData.length <= 0) return;

  sheet
    .getRange(data_height + 1, 1, newData.length, newData[0].length)
    .setValues(newData);
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
  return messages.flat().reverse();
}
