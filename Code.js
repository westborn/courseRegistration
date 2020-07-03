// GLOBAL constants for U3A

var U3A = {
  WORDPRESS_PROGRAM_FILE_ID: "1svCAoJKW7FsnerJSPhLkzuXEcicdksA5fcV2UfaztR8" // file is - "U3A Current Program - Wordpress"
};


// [START apps_script_menu]
/**
 * Handler for when a user opens the spreadsheet.
 * Creates a custom menu.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('U3A Menu')
    .addItem('First item', 'menuItem1')
    .addSeparator()
    .addSubMenu(ui.createMenu('WordPress Actions')
      .addItem('Update Course Program', 'makeCourseDetailForWordPress'))
    .addToUi();
}

/**
 * Handler for when menu item 2 is clicked.
 */
function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert('You clicked the second menu item!');
}
// [END apps_script_menu]



