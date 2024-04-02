function createMenuLinks() {
    // Create a custom menu
    const ui = SpreadsheetApp.getUi();

    let title = 'Project Management Settings';
    let scriptUser = getScriptUserEmail();
    if (scriptUser != null){
        title = `${title} (Sending as ${scriptUser})`
    }
    const menu = ui.createMenu(title);

    const userEmail = Session.getActiveUser().getEmail();

    if (userEmail != '') {
        saveUserEmail(userEmail);
    }

    let currentUser = getCustomUserEmail();

    if(currentUser == null) {
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

    if (getCustomUserEmail() != null) {
        createMenuLinks();
    }
    else {
        initiateMenu();
    }
    createButtonInCell();
}

function showDialog(page, title) {
    var html = HtmlService.createHtmlOutputFromFile(page)
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModalDialog(html, title);
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

function autoCreateTrigger() {
    deleteTriggerbyHandler('customEmailTrigger');
    createSpreadsheetOpenTrigger();
    createMenuLinks();

    const userEmail = Session.getActiveUser().getEmail();
    saveScriptUserEmail(userEmail);
    showDialog('TriggerInstall', 'Triggers Installed');
}
