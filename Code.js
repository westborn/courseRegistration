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
        .addItem('Import "Wordpress Enrolment" CSV', 'appendCSV')
        .addSeparator()
        .addItem('CSV', 'loadCSVSidebar')
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
 * Get a CSV file and write the transformed values to the "CSV" sheet
 *
 */
function appendCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName('CSV')
  const myFolder = getMyFolder(ss)
  const files = myFolder.searchFiles('Title contains "Wordpress Enrolments.csv"')
  const file = files.hasNext() ? files.next() : null

  if (!file) {
    const errText = 'No "Wordpress Enrolments.csv" file found in this folder'
    showToast(errText)
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 1).setValue(errText)
    Logger.log(errText)
    return
  }

  const csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ',')

  // get just the headers that we want (columns 5 -> second last)
  // this will be the original column sequence
  const headers = csvData.shift()
  const courseColumns = headers.slice(5, headers.length - 1)
  // check that the current form contains similar entries to the CSV
  // by matching the column count of each
  const currentColumns = sheet.getLastColumn()
  if (courseColumns.length !== currentColumns - 3) {
    const errText = 'CSV columns do not match current sheet'
    showToast(errText)
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 1).setValue(errText)
    Logger.log(errText)
    return
  }
  // sort alphabetic to create new column sequence
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

  //clear the sheet we are going to download the events to
  sheet.insertRowBefore(1)
  const lastRow = sheet.getLastRow()
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1)
  }
  //Write the data back to the sheet
  sheet.getRange(1, 1, result.length, result[0].length).setValues(result)

  //set a formula in the last column as error checking
  sheet.getRange(1, courseSequence.length + 3).setValue('errorCheck')
  for (let i = 0; i < result.length - 1; i++) {
    sheet
      .getRange(i + 2, courseSequence.length + 3)
      .setFormula(`ArrayFormula(index(Members,match(TRUE, exact(A${i + 2},memberName),0),1))`)
  }
}

function uploadFiles(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName('CSV')
  const errText = 'CSV loaded'
  showToast(errText)
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 1).setValue(errText)
  Logger.log(errText)
  return
}
