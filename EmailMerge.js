/**
 * simple loop to call "mergeDraftEmail" for every row in the database
 */
function allHTMLRegistrationEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const dbSheet = ss.getSheetByName('Database')
  let attendees = dbSheet.getRange('E13:E').getDisplayValues()
  const lastAttendeeIndex = attendees.filter(String).length
  // drop the last item - it is the Grand Total
  attendees.length = lastAttendeeIndex - 1
  attendees.forEach((attendee) => {
    mergeDraftEmail('TEMPLATE - Course Registration Information', {
      memberName: attendee[0],
      subject: 'U3A Bermagui - Course Registration Information',
    })
  })
}

/**
 * simple loop to call "mergeDraftEmail" for selected rows in the database
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
    mergeDraftEmail('TEMPLATE - Course Registration Information', {
      memberName: attendee[0],
      subject: 'U3A Bermagui - Course Registration Information',
    })
  })
}

/**
 * Get an existing draft temmplate and merge with a replacement object to produce a new draft email
 * @param {string} templateEmailSubject (optional) for the email draft template
 * @param {object} emailFields data fields for the new draft
 * @param {object} emailFields.memberName:
 * @param {object} emailFields.subject:
 * @param {object} emailFields.bcc:
 * @param {object} emailFields.cc:
 *
 * memberName:
 * firstName:
 * classDetails:
 */
function mergeDraftEmail(
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
    const msgObj = fillInTemplateFromObject_(emailTemplate.message, fieldReplacer)
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

  // // updating the sheet with new data
  // sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out)

  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
   */
  function getGmailTemplateFromDrafts_(subject_line) {
    try {
      // get drafts
      const drafts = GmailApp.getDrafts()
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0]
      // get the message object
      const msg = draft.getMessage()
      // getting attachments so they can be included in the merge
      const attachments = msg.getAttachments()
      return {
        message: { subject: subject_line, text: msg.getPlainBody(), html: msg.getBody() },
        attachments: attachments,
      }
    } catch (e) {
      throw new Error("Oops - can't find Gmail draft")
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
     */
    function subjectFilter_(subject_line) {
      return function (element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element
        }
      }
    }
  }

  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
   */
  function fillInTemplateFromObject_(template, data) {
    // we have two templates one for plain text and the html body
    // stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template)

    // token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, (key) => {
      return escapeData_(data[key.replace(/[{}]+/g, '')] || '')
    })
    return JSON.parse(template_string)
  }

  /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
   */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t')
  }
}
