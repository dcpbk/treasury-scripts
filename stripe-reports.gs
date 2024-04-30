const SHEET_LABEL = "auto-stripe-log";
const URL_CELL = "B1";

/**
 * Saves the target email label to the bound Spreadsheet.
 * @returns {void}
 */
function updateStripe() {
  // get the target sheet
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LABEL);
  const data_height = sheet.getDataRange().getHeight();

  // get the url for stripe transactions
  const reportURL = sheet.getRange(URL_CELL).getValue();

  // get ids for the transactions to filter new ones
  const transactionID = sheet
    .getRange(3, 1, data_height - 2, 1)
    .getValues()
    .flat();

  const newCSV = UrlFetchApp.fetch(reportURL);

  // read csv, ignoring header and existing transactions
  const newData = Utilities.parseCsv(newCSV.getContentText())
    .slice(1) // ignore header
    .filter((row) => !transactionID.includes(row[0])); // filter out existing

  // if new data is empty return
  if (newData.length <= 0) return;

  sheet
    .getRange(data_height + 1, 1, newData.length, newData[0].length)
    .setValues(newData);
}
