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
    .addItem('Import "Wordpress Enrolment" CSV', 'appendCSV')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('WordPress Actions')
        .addItem('Update Course Program', 'makeCourseDetailForWordPress')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Registration Advice Emails')
        .addItem('Draft ALL Registration Emails', 'allRegistrationEmails')
        .addItem('Draft SELECTED Registration Emails', 'selectedRegistrationEmails')    )  
    .addSubMenu(ui.createMenu('Zoom Actions').addItem('Schedule Zoom', 'selectedZoomSessions'))
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
  const file = DriveApp.getFilesByName('Wordpress Enrolments.csv').next()
  const csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ',')

  // get just the headers that we want (columns 5 -> second last)
  // this will be the original column sequence
  // sort alphabetic to create new column sequence
  const headers = csvData.shift()
  const courseColumns = headers.slice(5, headers.length - 1)
  const courseSequence = courseColumns.concat().sort()
  //loop thru CSV rows
  //  then thru columns of courses
  //    output name, email, [each course in alpha sequence]
  let result = [['name', 'email', ...courseSequence]]
  csvData.map((row) => {
    let thisRow = [row[3].trim(), row[4].trim(), ...Array.from({ length: courseSequence.length })]
    const courseCols = row.slice(5, row.length - 1)
    courseCols.map((col, idx) => {
      if (col != '') {
        const newCol = courseSequence.indexOf(courseColumns[idx])
        thisRow[newCol + 2] = '1'
      }
    })
    result.push(thisRow)
  })

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CSV')
  //clear the sheet we are going to download the events to
  sheet.insertRowBefore(1)
  const lastRow = sheet.getLastRow()
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1)
  }
  //Write the data back to the sheet
  sheet.getRange(1, 1, result.length, result[0].length).setValues(result)
 
  //set a formula in the last column as error checking 
  sheet.getRange(1,courseSequence.length+3).setValue("errorCheck")
  for (let i = 0; i < result.length-1; i++) {
    sheet.getRange(i+2, courseSequence.length+3).setFormula(`vlookup(A${i+2},memberName,1,false)`)
  }
}





