var SCRIPT_NAME = 'Email_Generator_Bound_Script',
	SCRIPT_VERSION = 'v1.9';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('CloudFire')
    .addItem('Create HTML', 'showMailPopup')
    .addItem('Set Defaults', 'showFormPopup')
    .addSeparator()
    .addItem('Delete Columns', 'deleteColumns')
    .addToUi();
}

// Menu
function showMailPopup() {EmailGenerator.showMailPopup()}
function showFormPopup() {EmailGenerator.showFormPopup()}
function deleteColumns() {EmailGenerator.deleteColumns()}

// Client-Side
function setDefaultValues(values) {EmailGenerator.setDefaultValues(values)}