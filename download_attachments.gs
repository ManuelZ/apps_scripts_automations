
/**
 * Convert a Date object to a formatted string in the format `yyyy_mm_dd`.
 * 
 * @param {Date} date - The date to format.
 * @returns {string} The formatted date string.
 */
function date_to_str(date) {
  var dd = date.getDate().toString().padStart(2, '0');
  var mm = (date.getMonth() + 1).toString().padStart(2, '0');  // Month from 0 to 11
  var yyyy = date.getFullYear();
  return `${yyyy}_${mm}_${dd}`;
}


/**
 * Download email attachments from a specific sender and saves them in a Google Drive folder.
 * Attachments are named with the email's received date and their original file name.
 *
 * Global constants:
 * - `OUTPUT_FOLDER_ID`: The ID of the Google Drive folder where attachments are saved.
 * - `SENDER_EMAIL`: The email address of the sender whose attachments will be downloaded.
 * 
 * - Emails from the sender are searched using Gmail search syntax.
 * - Attachments are saved in the specified folder with a date prefix in their file names.
 */
function download_attachments_from_sender() {
  const OUTPUT_FOLDER_ID = "";
  const SENDER_EMAIL = "";

  var threads = GmailApp.search('from:' + SENDER_EMAIL);

  threads.forEach(function (thread) {
    try {
      var messages = thread.getMessages();
      messages.forEach(function (message) {

        // Date when the email was received
        var date = message.getDate();
        var date_str = date_to_str(date);

        var attachments = message.getAttachments();
        attachments.forEach(function (attachment) {
          try {
            var folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
            var attachment_name = date_str + '_' + attachment.getName();
            var file = folder.createFile(attachment.copyBlob());
            file.setName(attachment_name)
            Logger.log('Saved attachment: ' + file.getName());
          } catch (e) {
            Logger.log('Error saving attachment: ' + e.toString());
          }
        });
      });

    } catch (e) {
      Logger.log('Error processing thread: ' + e.toString());
    }
  });
}



