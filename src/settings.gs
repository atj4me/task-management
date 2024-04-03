function createMenuLinks() {
    // Create a custom menu
    const ui = SpreadsheetApp.getUi();

    let title = 'Project Management Settings';
    let scriptUser = getScriptUserEmail();

    if (scriptUser != null && scriptUser != '') {
        title = `${title} (Sending as ${scriptUser})`;
    }
    const menu = ui.createMenu(title);

    const userEmail = Session.getActiveUser().getEmail();

    if (userEmail != '') {
        saveUserEmail(userEmail);
    }

    let currentUser = getCustomUserEmail();

    if (currentUser == null || currentUser == '') {
        currentUser = 'the current user';
    }

    // Add a button to the menu

    menu.addItem(`Send the e-mail as ${currentUser}`, 'autoCreateTrigger');
    menu.addItem('Create Sent Mail Dropdowns', 'createButtonInCell');

    // menu.addItem('Delete All Triggers (DO NOT USE)', 'deleteEmailTriggers');

    menu.addToUi();
}

function initiateMenu() {
    // Create a custom menu
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Project Management Settings');

    // Add a button to the menu
    menu.addItem(`Initiate Scripts`, 'createMenuLinks');
    menu.addToUi();
}

function onOpen() {

    let currentUser = getCustomUserEmail();
    if (currentUser != null && currentUser != '') {
        createMenuLinks();
    }
    else {
        initiateMenu();
    }
    createButtonInCell();
}

function deleteEmailTriggers() {
    deleteTriggerbyHandler('customEmailTrigger');
    createMenuLinks();
    showDialog('TriggerRemove', 'Triggers Removed');
}

function createSpreadsheetOpenTrigger() {
    const ss = SpreadsheetApp.getActive();
    let trigger = ScriptApp.newTrigger('customEmailTrigger')
        .forSpreadsheet(ss)
        .onEdit()
        .create();
    savetriggerId(trigger.getUniqueId());
    return trigger;
}

/**
 * Deletes a trigger.
 * @param {string} triggerId The Trigger ID.
 * @see https://developers.google.com/apps-script/guides/triggers/installable
 */
function deleteTrigger(triggerId) {
    // Loop over all triggers.
    const allTriggers = ScriptApp.getProjectTriggers();
    for (let index = 0; index < allTriggers.length; index++) {
        // If the current trigger is the correct one, delete it.
        if (allTriggers[index].getUniqueId() === triggerId) {
            ScriptApp.deleteTrigger(allTriggers[index]);
            break;
        }
    }
}

function autoCreateTrigger() {
    deleteTriggerbyHandler('customEmailTrigger');
    createSpreadsheetOpenTrigger();
    createMenuLinks();

    const userEmail = Session.getActiveUser().getEmail();
    saveScriptUserEmail(userEmail);
    showDialog('TriggerInstall', 'Triggers Installed');
}
