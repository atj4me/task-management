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
        .createHtmlOutput(`<div style="background: #FFFF00">${msg}</div>`)
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