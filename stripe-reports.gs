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
  const oldData = sheet.getDataRange().getValues();
  const data_height = oldData.length;

  // get the url for stripe transactions
  const reportURL = sheet.getRange(URL_CELL).getValue();

  // get ids for the transactions to filter new ones
  const transactionID = oldData.slice(2).map((row) => row[0]);

  const newCSV = UrlFetchApp.fetch(reportURL);

  // define array for updating bank deposit dates
  const update_automatic_payout_effective = [];

  // read csv and ignore rows that we already have
  const newData = Utilities.parseCsv(newCSV.getContentText())
    .slice(1)
    .filter((row) => {
      // store the index of the row in the old data (-2)
      const thisRowIndex = transactionID.indexOf(row[0]);

      // if the row is not in the sheet, the index will be -1
      // in this case, we want to include it in the new data
      if (thisRowIndex === -1) {
        return true;
      } else if (!oldData[thisRowIndex + 2][2] && row[2]) {
        // if the row is in the sheet, but the bank deposit date is empty
        // and the new row has a bank deposit date, store the row index and the new bank deposit date
        update_automatic_payout_effective.push([thisRowIndex, row[2]]);
      }

      // if the row is already in the sheet, exclude it
      return false;
    });

  // if new data is nonempty, add the new values
  if (newData.length > 0) {
    sheet
      .getRange(data_height + 1, 1, newData.length, newData[0].length)
      .setValues(newData);
  }

  // if there are new bank deposit dates, update them
  if (update_automatic_payout_effective.length > 0) {
    update_automatic_payout_effective.forEach((row) =>
      sheet.getRange(row[0] + 3, 3).setValue(row[1])
    );
  }
}
