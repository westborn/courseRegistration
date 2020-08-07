/**
 * Create an attendance register for signup on Information & Registration day
 * NOTE: This uses a Google Doc as a template for the attendance register.
 *       The docId for the template is '1LF4g60AxrZoJnsxqS2Gat_9UCIhkfeHIyTlxeOBfZR8'
 *
 * @param {string} templateDocId default '1LF4g60AxrZoJnsxqS2Gat_9UCIhkfeHIyTlxeOBfZR8'
 * @param {string} courseSummary of an existing course
 *
 */
function createAttendanceRegister(
  templateDocId = '1LF4g60AxrZoJnsxqS2Gat_9UCIhkfeHIyTlxeOBfZR8',
  courseSummary = 'How to buy wine you like ONLINE with Bhagya'
) {
  //get courseDetail sheet
  const courseData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('CourseDetails')
    .getDataRange()
    .getValues()
  const allCourses = getJsonArrayFromData(courseData)

  const thisCourse = allCourses.find(
    (course) => course.summary.toString().toLowerCase() === courseSummary.toString().toLowerCase()
  )

  const fileName = thisCourse.summary
  const titleDetails = thisCourse.summary
  const startDetails = `${thisCourse.days} ${thisCourse.dates} ${thisCourse.time}`
  const locationDetails = thisCourse.location
  const maxDetails = thisCourse.max.toString()
  const contactDetails = `${thisCourse.contact} ${thisCourse.phone}`

  //get memberDetail sheet
  const memberData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('MemberDetails')
    .getDataRange()
    .getValues()
  const allMembers = getJsonArrayFromData(memberData)

  const allMemberText = allMembers.map((member) => `${member.surname}\t${member.firstName}\n`)

  const myFolder = getMyFolder(SpreadsheetApp.getActiveSpreadsheet())
  const folder = checkIfFolderExistElseCreate(myFolder, 'Registration Lists')

  const fileList = folder.getFilesByName(fileName)
  if (fileList.hasNext()) {
    fileList.next().setTrashed(true)
  }
  registrationListFile = DriveApp.getFileById(templateDocId).makeCopy(fileName, folder)

  const replaceTextObject = {
    '{{titleDetails}}': titleDetails,
    '{{startDetails}}': startDetails,
    '{{locationDetails}}': locationDetails,
    '{{maxDetails}}': maxDetails,
    '{{contactDetails}}': contactDetails,
    '{{allMembers}}': allMemberText.sort().join(''),
  }

  findAndReplace(registrationListFile.getId(), replaceTextObject)
}

/**
 * Performs "replace all" in a Google Doc
 * @param {string} documentId The document to perform the replace text operations on.
 * @param {Object} findTextToReplacementMap A map from the "find text" to the "replace text".
 */
function findAndReplace(documentId, findTextToReplacementMap) {
  var requests = []
  for (var findText in findTextToReplacementMap) {
    var replaceText = findTextToReplacementMap[findText]
    var request = {
      replaceAllText: {
        containsText: {
          text: findText,
          matchCase: true,
        },
        replaceText: replaceText,
      },
    }
    requests.push(request)
  }

  var response = Docs.Documents.batchUpdate({ requests: requests }, documentId)
}
