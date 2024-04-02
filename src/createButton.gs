function createButtonInCell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cells = sheet.getRange("H8:H"); // Change to the desired cell address

  // Create a data validation rule
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      "1. Send Assignment e-mail",
      "2. Don't send e-mail",
      "3. Mail Sent",
      "4. Failed to Send E-mail"
    ])
    .setAllowInvalid(false)
    .build();

  // Apply the rule to the range
  cells.setDataValidation(rule);
}
