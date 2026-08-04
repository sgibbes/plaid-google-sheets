/** @format */

function testOnEdit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const commandCell = "G3";

  if (!/^[A-Z]+[1-9]\d*$/i.test(commandCell)) {
    throw new Error(
      'Command cell must use A1 notation, such as "G3". Received: "' + commandCell + '".'
    );
  }

  const fakeRange = sheet.getRange(commandCell);

  // Simulate the edit
  fakeRange.setValue(true);

  const fakeEvent = {
    range: fakeRange,
    value: "TRUE",
    source: SpreadsheetApp.getActiveSpreadsheet(),
    oldValue: "FALSE", // optional
    authMode: ScriptApp.AuthMode.FULL,
  };

  handleEdit_(fakeEvent);
}
