var SCRIPT_NAME = "Planning_Calendars_Collated_Bound_Script";
var SCRIPT_VERSION = "v0.2";

function onInstall() {
  onOpen()
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CloudFire')
      .addItem("Import 'Promotion Planning Calendars for Teams' data to 'Collated Data' sheet", "collatePlanningMaterial")
      .addItem("Sort 'Collated Data' by Column D", "sortCollatedData")
      .addSeparator()
      .addItem("Export 'Collated Data' to 'Promotion Deadlines Calendar'", "exportData")
      .addToUi();
}

// Menu items
function collatePlanningMaterial() {PCC.collatePlanningMaterial()}
function sortCollatedData()        {PCC.sortCollatedData()}
function exportData()              {PCC.exportData()}

// Triggers
function onEdit() {PCC.onEdit()}