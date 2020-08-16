/**
 * Get a list of all the users Calendars
 * @returns {object} containing the summary and id of each calendar found
 */
function getCalendarList() {
  var calendars
  let pageToken
  const result = []
  do {
    calendars = Calendar.CalendarList.list({
      maxResults: 100,
      pageToken: pageToken,
    })
    if (calendars.items && calendars.items.length > 0) {
      for (let i = 0; i < calendars.items.length; i++) {
        const calendar = calendars.items[i]
        result.push({ summary: calendar.summary, id: calendar.id })
        //        console.log('%s (ID: %s)', calendar.summary, calendar.id)
      }
    } else {
      console.log('No calendars found.')
    }
    pageToken = calendars.nextPageToken
  } while (pageToken)
  return result
}

/**
 * get all the events for the term from a calendar and write them to the "Calendar Download" sheet
 * @param {number} term number (1 - 4)
 * @param {id} calendarId of the calendar to retrieve events from
 */
function downloadCalendarEvents({ term = 3, calendarId = 'u3acomputerclub@hotmail.com' } = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheetOptions = ss.getSheetByName('Options')
  const sheetDownload = ss.getSheetByName('Calendar Download')

  // setup dates from the term selected (start and end dates)
  const terms = sheetOptions.getRange(1, 1, 5, 3).getValues()
  const eventRequest = {
    singleEvents: false,
    timeMin: new Date(terms[term][1]).toISOString(),
    timeMax: new Date(terms[term][2]).toISOString(),
  }

  //  get the events and format them
  const courseEvents = retrieveCalendarEvents(calendarId, eventRequest)

  //clear the sheet we are going to download the events to
  sheetDownload.insertRowBefore(2)
  const lastRow = sheetDownload.getLastRow()
  if (lastRow > 2) {
    sheetDownload.deleteRows(3, lastRow - 2)
  }

  if (!courseEvents.length) {
    sheetDownload.getRange(2, 1).setValue('No events Found')
    return
  }

  // drop all the recurrent event types (maybe used later?)
  const filteredEvents = courseEvents.filter((event) => event.type != '1-recurrent')

  //for each course - loop thru all sessions of the same name and resolve days/dates the course is scheduled.

  sheetEvents = filteredEvents.map((course) => {
    const thisSession = filteredEvents.filter((session) => session.summary === course.summary)
    const days = thisSession.map((el) => new Date(el.eventStartDateTime))
    let dedupDates = Array.from(new Set(days))
    dedupDates.sort((a, b) => {
      return new Date(a) - new Date(b)
    })
    course.datesScheduled = dedupDates
      .map((el) =>
        new Date(el)
          .toLocaleString('en-AU', {
            month: 'short',
            day: 'numeric',
          })
          .replace(' ', '-')
      )
      .join(', ')
    const dedupDays = Array.from(new Set(dedupDates.map((el) => new Date(el).getDay()))).sort()
    course.daysScheduled = dedupDays
      .map((el) => ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][el])
      .join(', ')
    return course
  })

  const rows = sheetEvents.map((d) => flatten_(d))
  const heads = sheetDownload.getDataRange().offset(0, 0, 1).getValues()[0]

  // convert object data into a 2d array
  const tr = rows.map((row) => heads.map((key) => row[String(key)] || ''))

  // write result
  sheetDownload.getRange(sheetDownload.getLastRow() + 1, 1, tr.length, tr[0].length).setValues(tr)

  return
}

/**
 * find an embedded contact name in a string
 * @param {string} description the string to be searched
 * @returns {string} contact name, if found
 */
const decodeContact = (description) => {
  const searchForContact = description.indexOf('Contact:')
  if (searchForContact > 0) {
    return description
      .slice(searchForContact + 9)
      .trim()
      .replace('.', '')
  } else {
    return ''
  }
}
/**
 * find an embedded presenter name in a string
 * @param {string} description the string to be searched
 * @returns {string} presenter name, if found
 */
const decodePresenter = (summary) => {
  const searchForPresenter = summary.match(/with(?!.*with)/i)
  if (searchForPresenter && searchForPresenter.index) {
    return summary.slice(searchForPresenter.index + 5).trim()
  } else {
    return ''
  }
}

/**
 * Extract course events (dates, location, summary, description) as an array of objects
 * @param {id} calendarId of the calendar to be extracted
 * @param {object} eventRequest containing parametres for the calendar search (type of search, start date/time)
 * @returns {object} courseEvent
 */
function retrieveCalendarEvents(calendarId, eventRequest) {
  const unpackEvent = (type, event) => {
    const courseEvent = {
      summary: event.summary || '',
      description: event.description ? stripHTML(event.description) : '',
      location: event.location || '',
      startDateTime: googleSheetDateTime(event.start.dateTime),
      endDateTime: googleSheetDateTime(event.end.dateTime),
      duration: '',
      daysScheduled: '',
      datesScheduled: '',
      presenter: '',
      contact: '',
      type: type,
      id: event.id || '',
      eventStartDateTime: event.start.dateTime,
      eventEndDateTime: event.end.dateTime,
    }
    const duration = dateDiffMinutes(new Date(event.end.dateTime), new Date(event.start.dateTime))
    courseEvent.duration = getTextTime(duration)
    courseEvent.presenter = decodePresenter(courseEvent.summary)
    courseEvent.contact = decodeContact(courseEvent.description)

    return courseEvent
  }

  const courseEvents = []

  eventRequest.singleEvents = false
  const singleEvents = Calendar.Events.list(calendarId, eventRequest).items.filter(
    (event) =>
      event &&
      event.hasOwnProperty('status') &&
      event.status !== 'cancelled' &&
      !event.start.hasOwnProperty('date') &&
      !(event.hasOwnProperty('recurrence') && event.recurrence.length > 1)
  )

  singleEvents.forEach((event) => {
    if (!event.recurrence) {
      if (event.recurringEventId) {
        courseEvents.push(unpackEvent('0-exception', event))
      } else courseEvents.push(unpackEvent('3-standalone', event))
    } else {
      courseEvents.push(unpackEvent('1-recurrent', event))
    }
  })

  const recurTypes = courseEvents.filter((event) => event.type === '1-recurrent')
  recurTypes.forEach((recur) => {
    const instanceEvents = Calendar.Events.instances(calendarId, recur.id).items.filter(
      (event) =>
        event &&
        event.hasOwnProperty('status') &&
        event.status !== 'cancelled' &&
        !event.start.hasOwnProperty('date') &&
        !(event.hasOwnProperty('recurrence') && event.recurrence.length > 1)
    )
    instanceEvents.forEach((el) => {
      const exists = courseEvents.find((obj) => obj.type === '0-exception' && obj.id === el.id)
      if (!exists) {
        courseEvents.push(unpackEvent('2-instance', el))
      }
    })
  })
  // console.log(courseEvents.length)
  courseEvents.sort((a, b) => new Date(a.eventStartDateTime) - new Date(b.eventStartDateTime))
  // courseEvents.forEach((e) => console.log(`${e.type} - ${e.id} - ${e.summary}`))
  //  console.log("\n\nException")
  //  exception.forEach(e => console.log(`${e.type} - ${e.id} - ${e.summary}`))
  //  console.log("\n\nRecurrent")
  //  recurrent.forEach(e => console.log(`${e.type} - ${e.id} - ${e.summary}`))

  return courseEvents
}

// Based on https://stackoverflow.com/a/54897035/1027723
const flatten_ = (obj, prefix = '', res = {}) =>
  Object.entries(obj).reduce((r, [key, val]) => {
    const k = `${prefix}${key}`
    if (typeof val === 'object' && val !== null) {
      flatten_(val, `${k}_`, r)
    } else {
      res[k] = val
    }
    return r
  }, res)
