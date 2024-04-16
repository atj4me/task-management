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


function showMessage(msg, type) {

    //Replace next lines with <br>
    msg.replace(/\n\r|\n|\r/g, "<br>")

    if (type == 'error') {
        msg = `<p style="color: red; font-weight: bold">${msg}</p>`;
        title = 'Error!';
    }
    else if (type == 'success') {
        msg = `<p style="color: green;">${msg}</p>`;
        title = 'Success!';
    }

    // Display a modal dialog box with custom HtmlService content.
    var htmlOutput = HtmlService
        .createHtmlOutput(`<body style="background: #FFFF00; padding: 1em 2em;">${msg}</body>`)
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModalDialog(htmlOutput, title);
}

function sendEmail(to, name, subject, body) {

    // Replace #Name with the actual name from cell B
    body = body.replace('#Name', name);

    // Send the email
    MailApp.sendEmail({
        to: to,
        subject: subject,
        body: body,
    });

    // Get the current timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    // Get the spreadsheet and sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("E-mail Log");

    // Append a new row with email details
    sheet.appendRow([
        name, // Sender Name (optional)
        to,
        subject,
        body,
        timestamp,
        getScriptUserEmail()
    ]);

    body = body.replace(/\n\r|\n|\r/g, "<br>");
    showMessage(`Mail has been Sent Successfully to ${name}<${to}>!  with the following content <br><br><pre>${body}</pre>`,
        'success');
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
        const formattedDate = Utilities.formatDate(dateString, 'en_US', 'EEEE, MMMM dd yyyy');
        return formattedDate;
    }
}

function showDialog(page, title) {
    var html = HtmlService.createHtmlOutputFromFile(`templates/${page}`)
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModalDialog(html, title);
}

/**
 * Deleted Trigger by Handler
 */
function deleteTriggerbyHandler(handler) {
    // Loop over all triggers.
    const allTriggers = ScriptApp.getProjectTriggers();
    for (let index = 0; index < allTriggers.length; index++) {
        // If the current trigger is the correct one, delete it.
        if (allTriggers[index].getHandlerFunction() === handler) {
            ScriptApp.deleteTrigger(allTriggers[index]);
        }
    }
    savetriggerId('');
}

function savetriggerId(trigger) {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('TRIGGER_ID', trigger);
}

function getTriggerId() {
    try {
        // Get the value for the user property 'DISPLAY_UNITS'.
        const scriptProperties = PropertiesService.getScriptProperties();
        const trigger = scriptProperties.getProperty('TRIGGER_ID');
        return trigger;
    } catch (err) {
        // TODO (developer) - Handle exception
        console.log('Failed with error %s', err.message);
    }
}

function saveUserEmail(email) {
    const scriptProperties = PropertiesService.getUserProperties();
    scriptProperties.setProperty('USER_EMAIL', email);
}

function getCustomUserEmail() {
    try {
        // Get the value for the user property 'DISPLAY_UNITS'.
        const scriptProperties = PropertiesService.getUserProperties();
        const trigger = scriptProperties.getProperty('USER_EMAIL');
        return trigger;
    } catch (err) {
        // TODO (developer) - Handle exception
        console.log('Failed with error %s', err.message);
    }
}

function saveScriptUserEmail(email) {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('USER_EMAIL', email);
}

function getScriptUserEmail() {
    try {
        // Get the value for the user property 'DISPLAY_UNITS'.
        const scriptProperties = PropertiesService.getScriptProperties();
        const trigger = scriptProperties.getProperty('USER_EMAIL');
        return trigger;
    } catch (err) {
        // TODO (developer) - Handle exception
        console.log('Failed with error %s', err.message);
    }
}

function checkTaskDelayed() {
    // Triggered when a cell is edited
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");

    if (sheet.getName() !== "Database") {
        return; // Exit if not the desired sheet
    }

    // Daily check: today's date in YYYY-MM-DD format
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Get all data at once (improves performance)
    const data = sheet.getDataRange().getValues();

    // Loop through each row of data
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const primaryHead = row[0];
        const tasks = row[1];
        const status = row[2]; // Assuming Column C is at index 2 (0-based indexing)
        const priority = row[3];
        const name = row[4];
        const due_date = row[6]; // Assuming Due Date is at index 7

        // Convert finish date and due date to proper date objects (assuming they are in YYYY-MM-DD format)
        const dueDateObj = new Date(due_date);
        dueDateObj.setHours(0, 0, 0, 0);

        if (status === "In Progress" || status == null) {
            // Perform check and update logic based on today and due date
            let newStatus = status;

            if (today.getTime() > dueDateObj.getTime()) {
                newStatus = "Delayed/Not completed";

                row[7] = ("1. Send Assignment e-mail");

                // Send E-mail when the task is delayed

                // Get data from the sheet
                const subject = primaryHead + ' task delayed - ' + priority + ' Priority';
                let emailBody = 'Hello #Name, \n' +
                    'This is to let you know that the following task has been delayed. \n\n' +
                    'Task Name: ' + tasks + '\n' +
                    'Status: ' + newStatus + '\n' +
                    'Priority: ' + priority + '\n' +
                    'Due Date: ' + formatDate(due_date) + '\n' +
                    '\n' +
                    'Thank You!';

                let recipientEmail = getEmailByName(name);

                // Validate recipient email (you can add more validation if needed)
                if (!isValidEmail(recipientEmail)) {
                    showMessage(
                        'The email is invalid!        ',
                        'error');
                    row[7] = ("4. Failed to Send E-mail");

                    return;
                }

                if (!tasks) {
                    showMessage(
                        'Please enter a valid task!        ',
                        'error');
                    row[7] = ("4. Failed to Send E-mail");

                    return;
                }

                if (!name) {
                    showMessage(
                        'Please select a valid Name!        ',
                        'error');
                    row[7] = ("4. Failed to Send E-mail");

                    return;
                }

                /**
                 * Function to Send Email
                 */
                sendEmail(recipientEmail, name, subject, emailBody);
                row[7] = ("3. Mail Sent");
            }

            // Update status only if changed
            if (newStatus !== status) {
                row[2] = newStatus; // Update status in the data array
            }
        }
    }

    // Update the sheet data in one go (improves performance)
    sheet.getDataRange().setValues(data);
}