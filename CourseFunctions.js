/**
 * Get data from the "RegistrationMaster" and populate the "Database"
 * For all the columns that have a 1 - write the details to the database columns
 * Recalculate the 2 pivot tables after the database is written back to the sheet
 */
function buildDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const ssId = ss.getId()
  const ssName = ss.getName()
  const ssFolder = DriveApp.getFolderById(ssId).getParents().next()
  //  showToast("Id: " + ssId + "  Name: " + ssName + "  Folder: " + ssFolder.getName());

  const fromSheet = ss.getSheetByName('RegistrationMaster')
  const fromData = fromSheet.getDataRange().getValues()
  const headers = fromData.shift()
  const totals = fromData.pop()
  //drop name, email, and count columns to get CourseNames ONLY
  courseNames = headers.slice(2, headers.length - 1)
  //loop thru rows
  //  then thru columns of courses
  //    select those with a 1
  //    output name and course
  const result = fromData.reduce((res, row) => {
    const courseCols = row.slice(2, row.length - 1)
    const tmp = courseCols.reduce((acc, col, idx) => {
      if (col === 1) {
        return acc.concat([[row[0], courseNames[idx]]])
      }
      return acc
    }, [])
    if (tmp.length) {
      return res.concat(tmp)
    }
    return res
  }, [])

  //clear 2 Database columns of ALL data
  dbSheet = ss.getSheetByName('Database')
  dbSheet.getRange('B13:C').clear()
  // write the 2 columns to the sheet - starting at "B13"
  dbSheet.getRange(13, 2, result.length, 2).setValues(result)

  //Now create the 2 pivot tables (E12 and H12) from the Database

  const sourceRange = 'B12:C' + (result.length + 12).toString()
  const sourceData = dbSheet.getRange(sourceRange)

  const pivotTable1 = dbSheet.getRange('E12').createPivotTable(sourceData)
  const pivotValue1 = pivotTable1.addPivotValue(
    2,
    SpreadsheetApp.PivotTableSummarizeFunction.COUNTA
  )
  pivotValue1.setDisplayName('numberCourses')
  const pivotGroup1 = pivotTable1.addRowGroup(2)

  const pivotTable2 = dbSheet.getRange('H12').createPivotTable(sourceData)
  const pivotValue2 = pivotTable2.addPivotValue(
    3,
    SpreadsheetApp.PivotTableSummarizeFunction.COUNTA
  )
  pivotValue2.setDisplayName('numberAttendees')
  const pivotGroup2 = pivotTable2.addRowGroup(3)
}

/**
 * Get the attendees from the sheet and populate a hyperlink with mailto: bcc items
 * 2 Hyperlinks are coonstructed 1 for Outlook (with ; delimiter) and one for Mac (with , delimiter)
 */
function makeHyperlink() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName('Attendance')

  var emailData = sheet.getRange('E14:E60').getValues()
  //flatten array and remove dups and drop empty strings
  var noDups = [...new Set(emailData.flat())].filter(String)
  //  Logger.log(noDups)
  //  Logger.log(noDups.length)
  sheet.getRange('C5:C6').clearContent()

  var hyperValOutlook = `=HYPERLINK("mailto:noreply@gmail.com?bcc=${noDups.join(
    ';'
  )}","Outlook Link")`
  var hyperValMac = `=HYPERLINK("mailto:noreply@gmail.com?bcc=${noDups.join(',')}","MacMail Link")`

  sheet.getRange('C5').setValue(hyperValOutlook)
  sheet.getRange('C5').setShowHyperlink(true)

  sheet.getRange('C6').setValue(hyperValMac)
  sheet.getRange('C6').setShowHyperlink(true)
}

/**
 * Write a print area to a PDF for Attendance data on the sheet
 * "O3" for the recipient
 * "D10" for presenter
 * "D8" for course
 */
function print_attendance() {
  makeHyperlink()
  var rangeNameToPrint = 'print_area_attendance'

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
  var presenterName = sheetToExport.getRange('D10').getDisplayValue()
  var courseTitle = sheetToExport.getRange('D8').getDisplayValue()
  var fileName = presenterName + '-' + courseTitle + '.pdf'

  var pdfFile = makePDFfromRange(selectedRange, fileName, 'Attendance Sheets')

  var recipient = sheetToExport.getRange('O3').getDisplayValue()
  var subject = courseTitle + ' - Attendance Sheet'
  var body =
    '\n\nAttached is the registration sheet for the course.\n\nPlease let us know if there are any changes required.\n\n\nU3A Team'

  var resp = GmailApp.createDraft(recipient, subject, body, {
    attachments: [pdfFile.getAs(MimeType.PDF)],
    name: 'Automatic Emailer Script',
  })

  return
}

/**
 * Write a print area to a PDF for Course data on the sheet
 * "K4" for the recipient name
 * "K2" for the recipient email
 * "B3" for Salutation (like "Hi George")
 */
function print_courseRegister() {
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
    '\n\nYour registration details for this term are listed on the attached PDF.\n\nPlease let us know if there are any changes required.\n\n\nU3A Team'

  var resp = GmailApp.createDraft(recipient, subject, body, {
    attachments: [pdfFile.getAs(MimeType.PDF)],
    name: 'Automatic Emailer Script',
  })

  return
}

/**
 * Create a draft email
 * @param {string} recipient for the email draft message
 * @param {string} subject for the email
 * @param {string} body of the email to send
 * @param {File} file descriptor for the PDF attachment
 */
function createDraft(recipient, subject, body, file) {
  var resp = GmailApp.createDraft(recipient, subject, body, {
    attachments: [file.getAs(MimeType.PDF)],
    name: 'Automatic Emailer Script',
  })
  //  Logger.log(resp);
}

/**
 * Create a draft email about a Zoom session from the "Attendance" sheet
 * "D8" for the course
 * "D9" for the time
 * "O3" for the recipient (presenter)
 * "E14:E60" for the bcc recipients
 *
 */
function createDraftZoomEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName('Attendance')

  var htmlBody = HtmlService.createTemplateFromFile('zoomReminder')
  htmlBody.course = sheet.getRange('D8').getDisplayValue()
  htmlBody.datetime = sheet.getRange('D9').getDisplayValue()
  var email_html = htmlBody.evaluate().getContent()

  var recipient = sheet.getRange('O3').getDisplayValue()
  var subject = 'U3A: ' + htmlBody.course + ' - ' + htmlBody.datetime
  var body = subject

  var userEmails = sheet.getRange('E14:E60').getValues()
  //flatten array and remove dups and drop empty strings
  var bccEmails = [...new Set(userEmails.flat())].filter(String).join(',')

  resp = GmailApp.createDraft(recipient, subject, body, {
    htmlBody: email_html,
    bcc: bccEmails,
  })
  //  Logger.log(resp);
}

/**
 * Create a formatted sheet that is displayed natively by WorPress
 * the wordPress sheet is pre-existing and the ID is a global reference
 * the data comes from the "CourseDetails" sheet
 *
 */
function makeCourseDetailForWordPress() {
  var rng = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('CourseDetails')
    .getDataRange()
    .getDisplayValues()
  numCourses = rng.length - 1

  var ssDest = SpreadsheetApp.openById(U3A.WORDPRESS_PROGRAM_FILE_ID)

  var sheet = ssDest.getSheetByName('Web Program')
  maxRows = sheet.getMaxRows()
  if (maxRows > 1) {
    sheet.deleteRows(2, maxRows - 1)
  }
  sheet.insertRowsAfter(1, numCourses - 1)

  for (var i = 1; i <= numCourses; i++) {
    //Loop through each row
    var outputStart = sheet.getRange(i, 1)
    courseDetailToSheet(rng[i], outputStart)
  }
}

/**
 * Create a formatted row from the CourseDetals sheet
 * @param {object} course row from CourseDetails
 * @param {range} outputTo range to write to on the sheet
 *
 * currently uses ordinal positions of the columns
 * TODO - use column headings OR pass in an object
 */
function courseDetailToSheet(course, outputTo) {
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build()
  var defaultFontSize = SpreadsheetApp.newTextStyle().setFontSize(12).build()
  var bodyFontSize = SpreadsheetApp.newTextStyle().setFontSize(11).build()
  var headColor = SpreadsheetApp.newTextStyle().setForegroundColor('#ff9900').build()
  var headFontSize = SpreadsheetApp.newTextStyle().setFontSize(14).build()

  let rich
  let cell

  cell = course[0] + '\n' + course[8]
  var headLen = course[0].length
  rich = SpreadsheetApp.newRichTextValue()
  rich
    .setText(cell)
    .setTextStyle(bodyFontSize)
    .setTextStyle(0, headLen, headColor)
    .setTextStyle(0, headLen, headFontSize)
  outputTo
    .setRichTextValue(rich.build())
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('middle')

  cell = course[5] + '\n' + course[6] + '\n' + course[7] + '\n' + course[12]
  rich = SpreadsheetApp.newRichTextValue()
  rich.setText(cell).setTextStyle(defaultFontSize)
  outputTo
    .offset(0, 1)
    .setRichTextValue(rich.build())
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('middle')

  outputTo
    .offset(0, 0, 1, 2)
    .setBorder(true, null, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID)
}

/**
 * simple loop to call "print_courseRegister" for every row in the database
 */
function allRegistrationEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const pdfSheet = ss.getSheetByName('Course Registration')
  const dbSheet = ss.getSheetByName('Database')
  let attendees = dbSheet.getRange('E13:E').getDisplayValues()
  const lastAttendeeIndex = attendees.filter(String).length
  // drop the last item - it is the Grand Total
  attendees.length = lastAttendeeIndex - 1

  attendees.forEach((attendee) => {
    //push the first name into the PDF sheet
    pdfSheet.getRange('K1').setValue(attendee[0])
    print_courseRegister()
  })
}

/**
 * simple loop to call "print_courseRegister" for selected rows in the database
 */
function selectedRegistrationEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const pdfSheet = ss.getSheetByName('Course Registration')
  const dbSheet = ss.getSheetByName('Database')

  const selectedRange = SpreadsheetApp.getActiveSpreadsheet().getActiveRange()
  selectedRange.activate()
  const selection = dbSheet.getSelection()
  const firstColumn = selection.getActiveRange().getColumn()
  const lastColumn = selection.getActiveRange().getLastColumn()

  // Must select one column only and must be column "E" (5)
  if (firstColumn != lastColumn || firstColumn != 5) {
    showToast('You need to Select one/some Member Names on the DATABASE sheet', 20)
    return
  }

  let attendees = selectedRange.getDisplayValues()
  attendees.forEach((attendee) => {
    //push the first name into the PDF sheet
    pdfSheet.getRange('K1').setValue(attendee[0])
    print_courseRegister()
  })
}
