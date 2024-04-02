function customEmailTrigger(e) {
    // Triggered when a cell is edited
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    if (sheet.getName() === "Database" && range.getColumn() === 8) {
        // Check if the edited cell is in column H (8)
        var value = range.getValue();
        if (value === "Send Assignment e-mail") {
            // Perform actions when "Yes" is selected (e.g., reset to "No")

            var row = range.getRow();

            // Update the corresponding cell in Column I
            var columnI = sheet.getRange('I' + row);

            // Get data from the sheet
            const tasks = sheet.getRange('B' + row).getValue(); // Change to the desired cell address
            const status = sheet.getRange('C' + row).getValue(); // Change to the desired cell address
            const priority = sheet.getRange('D' + row).getValue(); // Change to the desired cell address
            const due_date = sheet.getRange('G' + row).getValue(); // Change to the desired cell address
            const subject = 'A New Task has been assigned to you';
            let emailBody = 'Hello #Name, \n\r'+
                            'A new task has been assigned to you. Please find the details below \n\r \n\r'+
                            'Task Name: '+ tasks + '\n\r'+
                            'Status: ' + status + '\n\r'+
                            'Priority: ' + priority + '\n\r'+
                            'Due Date: ' + due_date + '\n\r'+
                            '\n\r' +
                            'Thank You!';

            // Replace #Name with the actual name from cell B
            const name = sheet.getRange('E' + row).getValue();
            emailBody = emailBody.replace('#Name', name);

            let recipientEmail = getEmailByName(name);
            
            // Validate recipient email (you can add more validation if needed)
            if (!isValidEmail(recipientEmail)) {
                showMessage(columnI,
                    'The email is invalid!        ',
                    'error');
                range.setValue("");
                return;
            }

            if (!subject) {
                showMessage(columnI,
                    'Please enter a valid Subject!        ',
                    'error');
                range.setValue("");
                return;
            }

            if (!emailBody) {
                showMessage(columnI,
                    'Please enter a valid Body!        ',
                    'error');
                range.setValue("");
                return;
            }

            if (!name) {
                showMessage(columnI,
                    'Please enter a valid Name!        ',
                    'error');
                range.setValue("");
                return;
            }

            /**
             * Function to Send Email
             */
            sendEmail(recipientEmail, subject, emailBody);

            // Resize the column width to fit content
            columnI.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

            showMessage(columnI,
                'Mail has been Sent Successfully to ' + recipientEmail + '!  with the following content \n\r      ' + emailBody,
                'success');

            sheet.autoResizeColumn(9); // Adjust the column index if needed
            range.setValue("");
        }
    }
}

// Validate email address
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

function getEmailByName(name) {
  // Get the spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("E-mails");
  
  // Get data range (excluding headers)
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  // Skip the header row
  var startIndex = 1;
  
  // Loop through each row (except header)
  for (var i = startIndex; i < values.length; i++) {
    if (values[i][0] === name) { // Check if name matches
      return values[i][1]; // Return email if found
    }
  }
  
  // If not found, return an appropriate message (optional)
  return false;
}


function showMessage(Cell, msg, type) {
    Cell.setValue(msg);
    // Apply formatting to highlight the cell
    Cell.setFontWeight('bold'); // Make the text bold
    Cell.setBackground('#FFFF00'); // Set a yellow background color

    if (type == 'error') {
        columnI.setFontColor('red'); // Change font color to red
    }
    else if (type == 'success') {
        Cell.setFontColor('green'); // Change font color to red
    }
}

function sendEmail(to, subject, body) {
    // Send the email
    MailApp.sendEmail({
        to: to,
        subject: subject,
        body: body,
    });
}