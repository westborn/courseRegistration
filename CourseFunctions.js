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
}

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

function createDraft(recipient, subject, body, file) {
  var resp = GmailApp.createDraft(recipient, subject, body, {
    attachments: [file.getAs(MimeType.PDF)],
    name: 'Automatic Emailer Script',
  })
  //  Logger.log(resp);
}

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

// pass in a course row as an array
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
