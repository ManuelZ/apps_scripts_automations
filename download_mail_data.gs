/*
This script extracts email data from Gmail based on a search query and writes it to a Google Spreadsheet. 
It processes emails in batches and handles pagination automatically.
*/

/**
 * Counts unique values in an iterable.
 * @param {Iterable<any>} iterable - Any iterable object
 * @return {number} Number of unique values
 */
function count_unique(iterable) {
  return new Set(iterable).size;
}


/**
 * Get the total number of rows in a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to count rows from
 * @return {number} The position of the last row that has content. 
 */
function get_num_rows(sheet) {
  return sheet.getLastRow();
}


/**
 * Append a 2D array of data to the next empty row in the specified sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to append data to
 * @param {Array<Array<any>>} data - 2D array of values to append
 */
function append_to_spreadsheet(sheet, data) {
  // Find the next empty row
  var lastRow = sheet.getLastRow();
  var nextEmptyRow = lastRow + 1;

  sheet.getRange(nextEmptyRow, 1, data.length, data[0].length).setValues(data);
}

/**
 * Retrieve the first row of the sheet (typically headers).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to read from
 * @return {Array<any>} Array of values from the first row
 */
function get_first_row(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * Retrieve all values from the first column of the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to read from
 * @return {Array<any>} Flat array of values from the first column
 */
function get_first_column_values(sheet) {
  return sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
}

/**
 * Fetch email data from Gmail based on the specified search criteria.
 * @param {number} start - Starting index for pagination
 * @param {string} query - Gmail search query string
 * @param {number} num_results - Number of thread results to fetch (max 500)
 * @return {Array<Array<string>>} 2D array containing [thread_id, sender, subject, date]
 */
function get_mail_data(start, query, num_results) {

  Logger.log(`Starting lookup from thread ${start}`);

  var data = [];

  var threads = GmailApp.search(query, start, num_results);
  threads.forEach(function (thread) {
    var thread_id = thread.getId();
    var first_message_subject = thread.getFirstMessageSubject();

    var messages = thread.getMessages();
    messages.forEach(function (message) {
      var sender = message.getFrom();
      var date_str = Utilities.formatDate(message.getDate(), 'America/Lima', 'yyyy-MM-dd HH:mm:ss.SSS Z')
      data.push([thread_id, sender, first_message_subject, date_str])
    });

  });

  return data;
}


/**
 * Main execution function that processes Gmail data into a spreadsheet.
 * Opens the specified spreadsheet and iteratively fetches email data
 * using a search query like "before:2025/01/31". Processes emails in
 * batches of 500 until no more matching emails are found.
 */
function main() {
  const SHEET_ID = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";  // FILL ME!
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheets()[0];

  const QUERY = "before:YYYY/MM/DD";  // FILL ME!
  const NUM_RESULTS = 500;  // Max is 500

  while (true) {

    var first_col_values = get_first_column_values(sheet);
    var start = count_unique(first_col_values) - 1;  // Minus one because the header doesn't count
    var data = get_mail_data(start, QUERY, NUM_RESULTS);

    if (data.length == 0) {
      Logger.log("End reached!");
      break
    }

    append_to_spreadsheet(sheet, data);
  }

}