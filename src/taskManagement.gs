function customEmailTrigger(e) {
    // Triggered when a cell is edited
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    if (sheet.getName() === "Database") {
        const row = range.getRow();
        var value = range.getValue();

        // Update the corresponding cell in Column I
        var columnI = sheet.getRange('I' + row);

        const finishDate = sheet.getRange(row, 6).getValue(); // Adjust this number if your finish date is in a different column
        const tasks = sheet.getRange('B' + row).getValue(); // Change to the desired cell address
        const name = sheet.getRange('E' + row).getValue();
        const priority = sheet.getRange('D' + row).getValue(); // Change to the desired cell address

        if (finishDate == '' &&
            tasks != '' &&
            name != '' &&
            priority != '') {

            let priorityName = 'In Progress';
            sheet.getRange(row, 3).setValue(priorityName);
        }


        if (range.getColumn() == 4) {
            let due_date = getDueDate(range.getValue());
            let due_date_col = sheet.getRange('G' + row);

            if (due_date != '' && 
                tasks != '' ) {
                due_date_col.setValue(due_date);
            }
        }
        else if (range.getColumn() == 6) {

            // Get the values for finish date, due date, and current status
            const dueDate = sheet.getRange(row, 8).getValue(); // Adjust this number if your due date is in a different column
            const currentStatus = sheet.getRange(row, 3).getValue();

            // Update status based on finish date and due date
            let newStatus;
            if (finishDate === "") {
                newStatus = "In Progress"; // No finish date, stays "In progress"
            } else if (finishDate < dueDate) {
                newStatus = "Done before time";
            } else if (finishDate > dueDate && currentStatus !== "Done before time") { // Avoid changing from "Done before time"
                newStatus = "Done but late";

                // Send E-mail when the task is delayed

                // Get data from the sheet
                const status = sheet.getRange('C' + row).getValue(); // Change to the desired cell address
                const due_date = sheet.getRange('G' + row).getValue(); // Change to the desired cell address
                const subject = 'The task has been delayed';
                let emailBody = 'Hello #Name, \n' +
                    'This is to let you know that the following task has been delayed. \n\n' +
                    'Task Name: ' + tasks + '\n' +
                    'Status: ' + status + '\n' +
                    'Priority: ' + priority + '\n' +
                    'Due Date: ' + formatDate(due_date) + '\n' +
                    'Finished Date: ' + formatDate(finishDate) + '\n' +
                    '\n' +
                    'Thank You!';

                // Replace #Name with the actual name from cell B
                emailBody = emailBody.replace('#Name', name);

                let recipientEmail = getEmailByName(name);

                // Validate recipient email (you can add more validation if needed)
                if (!isValidEmail(recipientEmail)) {
                    showMessage(columnI,
                        'The email is invalid!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("Failed to Send E-mail");
                    range.setValue("");
                    return;
                }

                if (!tasks) {
                    showMessage(columnI,
                        'Please enter a valid task!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("Failed to Send E-mail");
                    range.setValue("");
                    return;
                }

                if (!status) {
                    showMessage(columnI,
                        'Please select a valid Status!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("Failed to Send E-mail");
                    range.setValue("");
                    return;
                }

                if (!name) {
                    showMessage(columnI,
                        'Please select a valid Name!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("Failed to Send E-mail");
                    range.setValue("");
                    return;
                }

                /**
                 * Function to Send Email
                 */
                sendEmail(recipientEmail, subject, emailBody);
                sheet.getRange(row, 8).setValue("Mail Sent");
            } else {
                newStatus = "Delayed/Not completed";
            }

            // Update the status column only if it's different
            if (newStatus !== currentStatus) {
                sheet.getRange(row, 3).setValue(newStatus);
            }
        }
        else if (range.getColumn() === 8) {
            // Check if the edited cell is in column H (8)
            if (value === "Send Assignment e-mail") {
                // Perform actions when "Yes" is selected (e.g., reset to "No")


                // Get data from the sheet
                const status = sheet.getRange('C' + row).getValue(); // Change to the desired cell address
                const due_date = sheet.getRange('G' + row).getValue(); // Change to the desired cell address
                const subject = 'A New Task has been assigned to you';
                let emailBody = 'Hello #Name, \n' +
                    'A new task has been assigned to you. Please find the details below \n\n' +
                    'Task Name: ' + tasks + '\n' +
                    'Status: ' + status + '\n' +
                    'Priority: ' + priority + '\n' +
                    'Due Date: ' + formatDate(due_date) + '\n' +
                    '\n' +
                    'Thank You!';

                // Replace #Name with the actual name from cell B
                emailBody = emailBody.replace('#Name', name);

                let recipientEmail = getEmailByName(name);

                // Validate recipient email (you can add more validation if needed)
                if (!isValidEmail(recipientEmail)) {
                    showMessage(columnI,
                        'The email is invalid!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("Failed to Send E-mail");
                    return;
                }

                if (!tasks) {
                    showMessage(columnI,
                        'Please enter a valid task!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("Failed to Send E-mail");
                    return;
                }

                if (!status) {
                    showMessage(columnI,
                        'Please select a valid Status!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("Failed to Send E-mail");
                    return;
                }

                if (!name) {
                    showMessage(columnI,
                        'Please select a valid Name!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("Failed to Send E-mail");
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
                sheet.getRange(row, 8).setValue("Mail Sent");
            }
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
        Cell.setFontColor('red'); // Change font color to red
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

function getDueDate(priority) {
    // Get the spreadsheet object
    const sheet = SpreadsheetApp.getActiveSheet();
    let dueDate = '';

    if (priority === "High") {
        dueDate = getFutureWeekday(2);
    } else if (priority === "Medium") {
        dueDate = getFutureWeekday(3);
    } else if (priority === "Low") {
        dueDate = getFutureWeekday(4);
    } else {
        // Handle invalid priority (optional: leave blank, set default date, etc.)
        dueDate = "";
    }
    return dueDate;
}

// Function to get a future weekday (excluding weekends)
function getFutureWeekday(offset) {
    const today = new Date();
    let futureDate = new Date(today.getTime() + (offset * 24 * 60 * 60 * 1000));

    // Adjust for weekends
    while (futureDate.getDay() === 0 || futureDate.getDay() === 6) {
        futureDate.setDate(futureDate.getDate() + 1);
    }

    return futureDate;
}

function formatDate(dateString) {
    // Check if the value is actually a date (optional)
    if (dateString instanceof Date) {
        // Use Utilities.formatDate for flexible formatting
        const formattedDate = Utilities.formatDate(dateString, Session.getScriptLocale(), 'EEEE, MMMM dd yyyy');
        return formattedDate;
    }
}