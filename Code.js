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
    .addItem('Import Calendar', 'loadCalendarSidebar')
    .addItem('Import "Wordpress Enrolment" CSV', 'appendCSV')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('WordPress Actions')
        .addItem('Update Course Program', 'makeCourseDetailForWordPress')
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Registration Advice Emails')
        .addItem('Draft ALL Registration Emails', 'allRegistrationEmails')
        .addItem('Draft SELECTED Registration Emails', 'selectedRegistrationEmails')
    )
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

function btn_buildDB() {
  buildDB()
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
  sheet.getRange(1, courseSequence.length + 3).setValue('errorCheck')
  for (let i = 0; i < result.length - 1; i++) {
    sheet
      .getRange(i + 2, courseSequence.length + 3)
      .setFormula(`vlookup(A${i + 2},memberName,1,false)`)
  }
}

function correct_courseRegister() {
  var rangeNameToPrint = 'print_area_courseRegister'

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var myNamedRanges = getNamedRangesA1(spreadsheet)
  if (myNamedRanges[rangeNameToPrint] === 'undefined') {
    showToast(
      "No print area found. Please define one 'print_area_????' named range using Data > Named ranges.",
      30
    )
    return
  }
  var selectedRange = spreadsheet.getRangeByName(myNamedRanges[rangeNameToPrint])
  var sheetToExport = selectedRange.getSheet()
  var memberName = sheetToExport.getRange('K4').getDisplayValue()
  var fileName = memberName + ' - Enrolment Information.pdf'

  var pdfFile = makePDFfromRange(selectedRange, fileName, 'Enrolment Information')

  var recipient = sheetToExport.getRange('K2').getDisplayValue()
  var subject = sheetToExport.getRange('K3').getDisplayValue()
  var body =
    sheetToExport.getRange('B3').getDisplayValue() +
    '\n\nAttached is an updated registration with, we hope, the correct details for the sessions you registered for.' +
    '\n\nSorry for the error in processing, we trust this is all OK now?' +
    '\n\nYour registration details for this term are listed on the attached PDF.\n\nPlease let us know if there are any changes required.\n\n\nU3A Team'

  var resp = GmailApp.createDraft(recipient, subject, body, {
    attachments: [pdfFile.getAs(MimeType.PDF)],
    name: 'Bermagui U3A',
  })

  return
}
