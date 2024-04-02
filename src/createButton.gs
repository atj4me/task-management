function createButtonInCell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cells = sheet.getRange("H8:H"); // Change to the desired cell address

  // Create a data validation rule
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Send Assignment e-mail", "Don't send e-mail", "Mail Sent", "Failed to Send E-mail"])
    .setAllowInvalid(false)
    .build();

  // Apply the rule to the range
  cells.setDataValidation(rule);
}
