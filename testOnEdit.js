/** @format */

function testOnEdit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const fakeRange = sheet.getRange("g4"); // Row 1, Column 15 (O1)
  // const fakeRange = sheet.getRange(1, 2); // Row 1, Column 15 (O1)

  // Simulate the edit
  fakeRange.setValue("TRUE"); // Optional: mimic actual edit

  const fakeEvent = {
    range: fakeRange,
    value: "TRUE",
    source: SpreadsheetApp.getActiveSpreadsheet(),
    oldValue: "FALSE", // optional
    authMode: ScriptApp.AuthMode.FULL,
  };

  onEdit(fakeEvent);
}
