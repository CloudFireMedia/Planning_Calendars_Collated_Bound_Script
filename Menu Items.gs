function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CloudFire')
      .addItem("Import 'Promotion Planning Calendars for Teams' data to 'Collated Data' sheet", "collatePlanningMaterial")
      .addItem("Sort 'Collated Data' by Column D", "sortCollatedData")
      .addSeparator()
      .addItem("Export 'Collated Data' to 'Promotion Deadlines Calendar'", "exportData")
      .addToUi();
}
