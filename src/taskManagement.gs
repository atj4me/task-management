function customEmailTrigger(e) {
    // Triggered when a cell is edited
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    if (sheet.getName() === "Database") {
        const row = range.getRow();
        var value = range.getValue();

        // Update the corresponding cell in Column I
        var columnI = sheet.getRange('I' + row);

        const primaryHead = sheet.getRange(row, 1).getValue(); // Adjust this number if your finish date is in a different column
        const finishDate = sheet.getRange(row, 6).getValue(); // Adjust this number if your finish date is in a different column
        const tasks = sheet.getRange('B' + row).getValue(); // Change to the desired cell address
        const name = sheet.getRange('E' + row).getValue();
        const priority = sheet.getRange('D' + row).getValue(); // Change to the desired cell address
        const due_date = sheet.getRange('G' + row).getValue(); // Change to the desired cell address

        if (finishDate == '' &&
            tasks != '' &&
            name != '' &&
            priority != '') {

            let priorityName = 'In Progress';
            sheet.getRange(row, 3).setValue(priorityName);
            // sheet.getRange(row, 8).setValue("");
        }


        if (range.getColumn() == 4) {
            let due_date = getDueDate(range.getValue());
            let due_date_col = sheet.getRange('G' + row);

            // if (due_date != '' &&
            //     tasks != '') {
            //     due_date_col.setValue(due_date);
            // }

            due_date_col.setValue(due_date);
        }
        else if (range.getColumn() == 6) {

            const currentStatus = sheet.getRange(row, 3).getValue();

            // Update status based on finish date and due date
            let newStatus = currentStatus;

            const dueDateObj = new Date(due_date);
            dueDateObj.setHours(0, 0, 0, 0);

            const finishDateObj = new Date(finishDate);
            finishDateObj.setHours(0, 0, 0, 0);
            if (finishDate === "") {
                newStatus = "In Progress"; // No finish date, stays "In progress"
            } else if (finishDateObj.getTime() < dueDateObj.getTime()) {
                newStatus = "Done before time";
            } else if (finishDateObj.getTime() == dueDateObj.getTime()) {
                newStatus = "Done";
            } else if (finishDateObj.getTime() > dueDateObj.getTime() && currentStatus !== "Done before time") { // Avoid changing from "Done before time"
                newStatus = "Done but late";

                sheet.getRange(row, 8).setValue("1. Send Assignment e-mail");

                // Send E-mail when the task is delayed

                // Get data from the sheet
                const subject = primaryHead + ' task delayed - ' + priority + ' Priority';
                let emailBody = 'Hello #Name, \n' +
                    'This is to let you know that the following task has been delayed. \n\n' +
                    'Task Name: ' + tasks + '\n' +
                    'Status: ' + newStatus + '\n' +
                    'Priority: ' + priority + '\n' +
                    'Due Date: ' + formatDate(due_date) + '\n' +
                    'Finished Date: ' + formatDate(finishDate) + '\n' +
                    '\n' +
                    'Thank You!';

                let recipientEmail = getEmailByName(name);

                // Validate recipient email (you can add more validation if needed)
                if (!isValidEmail(recipientEmail)) {
                    showMessage(
                        'The email is invalid!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("4. Failed to Send E-mail");
                    range.setValue("");
                    return;
                }

                if (!tasks) {
                    showMessage(
                        'Please enter a valid task!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("4. Failed to Send E-mail");
                    range.setValue("");
                    return;
                }

                if (!name) {
                    showMessage(
                        'Please select a valid Name!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("4. Failed to Send E-mail");
                    range.setValue("");
                    return;
                }

                /**
                 * Function to Send Email
                 */
                sendEmail(recipientEmail, name, subject, emailBody);
                sheet.getRange(row, 8).setValue("3. Mail Sent");
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
            if (value === "1. Send Assignment e-mail") {
                // Perform actions when "Yes" is selected (e.g., reset to "No")


                // Get data from the sheet
                const status = sheet.getRange('C' + row).getValue(); // Change to the desired cell address
                const subject = primaryHead + ' task assigned - ' + priority + ' Priority';
                let emailBody = 'Hello #Name, \n' +
                    'A new task has been assigned to you. Please find the details below \n\n' +
                    'Task Name: ' + tasks + '\n' +
                    'Status: ' + status + '\n' +
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
                    sheet.getRange(row, 8).setValue("4. Failed to Send E-mail");
                    return;
                }

                if (!tasks) {
                    showMessage(
                        'Please enter a valid task!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("4. Failed to Send E-mail");
                    return;
                }

                if (!status) {
                    showMessage(
                        'Please select a valid Status!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("4. Failed to Send E-mail");
                    return;
                }

                if (!name) {
                    showMessage(
                        'Please select a valid Name!        ',
                        'error');
                    sheet.getRange(row, 8).setValue("4. Failed to Send E-mail");
                    return;
                }

                /**
                 * Function to Send Email
                 */
                sendEmail(recipientEmail, name, subject, emailBody);

                // Resize the column width to fit content
                columnI.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

                sheet.autoResizeColumn(9); // Adjust the column index if needed
                sheet.getRange(row, 8).setValue("3. Mail Sent");
            }
        }
    }
}

