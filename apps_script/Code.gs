function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Project EDGE')
    .addItem('Refresh Odds', 'refreshOdds')
    .addItem('Test Setup', 'testSetup')
    .addToUi();
}

function testSetup() {
  SpreadsheetApp.getUi().alert('Project EDGE is connected and working.');
}

function refreshOdds() {
  SpreadsheetApp.getUi().alert('Next step: live odds code will be added now.');
}
