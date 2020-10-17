/**
 * simple loop to call "draftEnrolleeEmail" for every row in the database
 */
function allHTMLRegistrationEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const dbSheet = ss.getSheetByName('Database')
  let attendees = dbSheet.getRange('E13:E').getDisplayValues()
  const lastAttendeeIndex = attendees.filter(String).length
  // drop the last item - it is the Grand Total
  attendees.length = lastAttendeeIndex - 1
  attendees.forEach((attendee) => {
    draftEnrolleeEmail('TEMPLATE - Course Registration Information', {
      memberName: attendee[0],
      subject: 'U3A Bermagui - Course Registration Information',
    })
  })
}

/**
 * simple loop to call "draftEnrolleeEmail" for selected rows in the database
 */
function selectedHTMLRegistrationEmails() {
  // Must select from the Database sheet and must be in column "E" (5)
  const res = metaSelected('Database', 5)
  if (!res) {
    return
  }
  const { sheetSelected, rangeSelected } = res
  let attendees = sheetSelected.getRange(rangeSelected).getDisplayValues()
  attendees.forEach((attendee) => {
    draftEnrolleeEmail('TEMPLATE - Course Registration Information', {
      memberName: attendee[0],
      subject: 'U3A Bermagui - Course Registration Information',
    })
  })
}

/**
 * Get an existing draft temmplate and merge with a replacement object to produce an Enrollee Email
 * with details of all the courses a member is attanding
 * @param {string} templateEmailSubject (optional) for the email draft template
 * @param {object} emailFields data fields for the new draft
 * @param {object} emailFields.memberName
 * @param {object} emailFields.subject
 *
 */
function draftEnrolleeEmail(
  templateEmailSubject = 'TEMPLATE - Course Registration Information',
  emailFields
) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!templateEmailSubject) {
    templateEmailSubject = Browser.inputBox(
      'Mail Merge',
      'Type or copy/paste the subject line of the Gmail ' +
        'draft message you would like to mail merge with:',
      Browser.Buttons.OK_CANCEL
    )

    if (templateEmailSubject === 'cancel' || templateEmailSubject == '') {
      // if no subject line finish up
      return
    }
  }

  //get courseDetail sheet
  const courseData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('CourseDetails')
    .getDataRange()
    .getValues()
  const allCourses = getJsonArrayFromData(courseData)

  //get the Database of who is attending which course (columns B:C)
  const db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database')
  const dbData = db.getRange('B12:C' + db.getLastRow()).getValues()
  const allDB = getJsonArrayFromData(dbData)

  // filter the Database for just this members courses
  const memberIsGoingTo = allDB
    .filter(
      (dbEntry) =>
        dbEntry.memberName.toString().toLowerCase() ===
        emailFields.memberName.toString().toLowerCase()
    )
    .map((entry) => entry.goingTo)

  // get the courseDetails rows for all the courses the member is attending
  const classInfo = memberIsGoingTo
    .map((courseTitle) =>
      allCourses
        .filter(
          (course) => course.title.toString().toLowerCase() === courseTitle.toString().toLowerCase()
        )
        .map((cR) => {
          const tmp = `
          <br>
          <b>${cR.title}</b><font color="#606060"> with ${cR.presenter}</font>
          <br>&nbsp;&nbsp;&nbsp;&nbsp;When: ${cR.days} ${cR.dates}
          <br>&nbsp;&nbsp;&nbsp;&nbsp;Time: ${cR.time}
          <br>&nbsp;&nbsp;&nbsp;&nbsp;Where: ${cR.location}
          <br>&nbsp;&nbsp;&nbsp;&nbsp;Contact: ${cR.contact}<br>
          `
          return tmp
        })
    )
    .join('\n')

  //get memberDetail sheet
  const memberData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('MemberDetails')
    .getDataRange()
    .getValues()
  const allMembers = getJsonArrayFromData(memberData)

  //find this member in the MemberDetails
  const thisMember = allMembers.find(
    (member) =>
      emailFields.memberName.toString().toLowerCase() === member.memberName.toString().toLowerCase()
  )

  const fieldReplacer = {
    memberName: emailFields.memberName,
    firstName: thisMember.firstName,
    classInfo,
  }

  // get the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(templateEmailSubject)

  try {
    const msgObj = fillinTemplateFromObject(emailTemplate.message, fieldReplacer)
    const msgText = stripHTML(msgObj.text)
    GmailApp.createDraft(thisMember.email, emailFields.subject, msgText, {
      htmlBody: msgObj.html,
      // bcc: 'a.bbc@email.com',
      // cc: 'a.cc@email.com',
      // from: 'an.alias@email.com',
      // name: 'name of the sender',
      // replyTo: 'a.reply@email.com',
      attachments: emailTemplate.attachments,
    })
  } catch (e) {
    throw new Error("Oops - can't create new Gmail draft")
  }
}
