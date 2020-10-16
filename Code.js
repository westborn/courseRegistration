/**
 * GLOBAL constants for U3A
 * Change these to match the column names you are using for email
 * recepient addresses and email sent column.
 */
var U3A = {
  WORDPRESS_PROGRAM_FILE_ID: '1svCAoJKW7FsnerJSPhLkzuXEcicdksA5fcV2UfaztR8', // file is - "U3A Current Program - Wordpress"
}

/**
 * Creates the menu items for user to run scripts on drop-down.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('U3A Menu')
    .addSubMenu(
      ui
        .createMenu('CalendarImport')
        .addItem('Schedule Zoom Meeting', 'selectedZoomSessions')
        .addItem('Email Zoom Session Advice', 'createZoomSessionEmail')
        .addItem('Import Calendar', 'loadCalendarSidebar')
        .addItem('Create CourseDetails', 'createCourseDetails')
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Database')
        .addItem('Email ALL Enrollees - PDF', 'allRegistrationEmails')
        .addItem('Email ALL Enrollees - HTML', 'allHTMLRegistrationEmails')
        .addItem('Email SELECTED Enrollees - PDF', 'selectedRegistrationEmails')
        .addItem('Email SELECTED Enrollees - HTML', 'selectedHTMLRegistrationEmails')
        .addItem('Create Database', 'buildDB')
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Wordpress Actions')
        .addItem('Create Course Program', 'makeCourseDetailForWordPress')
        .addItem('Import Enrolment CSV', 'loadCSVSidebar')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Other Actions').addItem('I&R Enrolment Sheet', 'selectedAttendanceRegister')
    )
    .addSeparator()
    .addItem('Help', 'loadHelpSidebar')
    .addToUi()
}

/**
 * Handler  to load Calendar Sidebar.
 */
function loadCalendarSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('calendarSidebar').setTitle('U3A Tools')
  SpreadsheetApp.getUi().showSidebar(html)
}

/**
 * Handler  to load Help Sidebar.
 */
function loadHelpSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('HelpSidebar').setTitle('U3A Tools Help')
  SpreadsheetApp.getUi().showSidebar(html)
}

/**
 * Handler  to load CSV Uploader Sidebar.
 */
function loadCSVSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('uploadCSV').setTitle('CSV Upload')
  SpreadsheetApp.getUi().showSidebar(html)
}

function btn_makeHyperlink() {
  makeHyperlink()
}

function btn_print_attendance() {
  print_attendance()
}

function btn_createDraftZoomEmail() {
  createDraftZoomEmail()
}

function btn_print_courseRegister() {
  print_courseRegister()
}

/**
 * Take the contents of a CSV file and write the transformed values to the "CSV" sheet
 *
 */
function appendCSV(csvData, writeMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName('CSV')

  //get courseDetail sheet
  const courseData = ss.getSheetByName('CourseDetails').getDataRange().getValues()
  const allCourses = getJsonArrayFromData(courseData)
  //get just the header tags from  the Course Details sheet
  headers = allCourses.map((course) => course.tag)

  //turn the CSV into objects
  const csvArray = getJsonArrayFromData(csvData)

  //sort headers alphabetic to create correct column sequence
  const courseSequence = headers.sort()
  let result = []

  //if we're deleting all existing, add the headings back
  //clear the sheet we are going to download to
  if (writeMode === 'create') {
    result = [['name', 'email', ...courseSequence]]
    sheet.insertRowBefore(1)
    const lastRow = sheet.getLastRow()
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1)
    }
  }

  //loop thru CSV rows (use all entries for "create" else use "unread" entries)
  //  then thru columns of courses
  //    output name, email, [each course]
  csvArray.map((entry) => {
    if (writeMode === 'create' || (writeMode === 'append' && entry.Status === 'unread')) {
      const thisRow = courseSequence.map((col) => {
        return entry[col] ? '1' : ''
      })
      result.push([entry.Name.trim(), entry.Email.trim(), ...thisRow])
    }
  })

  //Write the data back to the sheet
  sheet.getRange(sheet.getLastRow() + 1, 1, result.length, result[0].length).setValues(result)

  //set a formula in the last 2 columns as error checking
  sheet.getRange(1, courseSequence.length + 3, 1, 2).setValues([['nameCheck', 'emailCheck']])
  const formulas = [
    'ArrayFormula(index(Members,match(TRUE, exact(A2,memberName),0),1))',
    'ArrayFormula(index(Members,match(TRUE, exact(B2,memberEmail),0),1))',
  ]
  sheet.getRange(2, courseSequence.length + 3, 1, 2).setFormulas([formulas])
  const fillDownRange = sheet.getRange(2, courseSequence.length + 3, sheet.getLastRow() - 1)
  sheet.getRange(2, courseSequence.length + 3, 1, 2).copyTo(fillDownRange)
}

function readCSV(data, writeMode) {
  var csvFile = Utilities.newBlob(data.bytes, data.mimeType, data.filename)
  const csvData = Utilities.parseCsv(csvFile.getDataAsString(), ',')

  appendCSV(csvData, writeMode)
  return
}
