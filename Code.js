// GLOBAL constants for U3A

var U3A = {
  WORDPRESS_PROGRAM_FILE_ID: '1svCAoJKW7FsnerJSPhLkzuXEcicdksA5fcV2UfaztR8', // file is - "U3A Current Program - Wordpress"
}

// [START apps_script_menu]
/**
 * Handler for when a user opens the spreadsheet.
 * Creates a custom menu.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('U3A Menu')
    .addItem('Import Calendar', 'loadCalendarSidebar')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('WordPress Actions')
        .addItem('Update Course Program', 'makeCourseDetailForWordPress')
    )
    .addSubMenu(ui.createMenu('Zoom Actions').addItem('Schedule Zoom', 'createZoomMeeting'))
    .addToUi()
}

/**
 * Handler  to load Calendar Sidebar.
 */
function loadCalendarSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('calendarSidebar').setTitle('U3A Tools')
  SpreadsheetApp.getUi().showSidebar(html)
}

function appendCSV() {
  var file = DriveApp.getFilesByName('test reg 2.csv').next()
  Logger.log(file.getName())
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ',')
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CSV')
  Logger.log(sheet.getName())
  var lastRow = sheet.getLastRow()
  Logger.log(lastRow)
  sheet.getRange(lastRow + 1, 1, csvData.length, csvData[0].length).setValues(csvData)
}
