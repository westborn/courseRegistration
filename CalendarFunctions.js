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
 * @returns {object} containing the summary and id of each calendar found
 *
 * uses rrule library to decode recurrence events
 */
function downloadCalendarEvents({ term, calendarId } = {}) {
  eval(
    UrlFetchApp.fetch(
      'https://jakubroztocil.github.io/rrule/dist/es5/rrule-tz.min.js'
    ).getContentText()
  )

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

  const filteredEvents = courseEvents.filter((event) => event.type != '1-recurrent')
  const rows = filteredEvents.map((d) => flatten_(d))
  const heads = sheetDownload.getDataRange().offset(0, 0, 1).getValues()[0]

  // convert object data into a 2d array
  const tr = rows.map((row) => heads.map((key) => row[String(key)] || ''))

  // write result
  sheetDownload.getRange(sheetDownload.getLastRow() + 1, 1, tr.length, tr[0].length).setValues(tr)

  return
}

/**
 * remove any html tags and non-breaking-space from a string
 * @param {string} description the string to be manipulated
 * @returns {string} the munged result
 */
const decodeDescription = (description) => {
  return description
    .replace(/(<([^>]+)>)/gi, '')
    .replace(/&nbsp;/g, ' ')
    .trim()
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
      type: type,
      id: event.id || '',
      summary: event.summary || '',
      description: event.description ? decodeDescription(event.description) : '',
      location: event.location || '',
      startDateTime: new Date(event.start.dateTime).toLocaleString().replace(',', ''),
      endDateTime: new Date(event.end.dateTime).toLocaleString().replace(',', ''),
      recurrence: '',
      recurText: '',
      recurDates: '',
      presenter: getNested(event, 'extendedProperties', 'private', 'presenter'),
      contact: getNested(event, 'extendedProperties', 'private', 'contact'),
      //      min: getNested(event, 'extendedProperties', 'private', 'min'),
      //      max: getNested(event, 'extendedProperties', 'private', 'max'),
      //      cost: getNested(event, 'extendedProperties', 'private', 'cost'),
      //      isZoom: 'N',
      //      zoomLink: '',
    }

    // get Presenter and Contact from the text of the calendar
    if (!courseEvent.presenter) {
      courseEvent.presenter = decodePresenter(courseEvent.summary)
    }

    if (!courseEvent.contact) {
      courseEvent.contact = decodeContact(courseEvent.description)
    }

    if (event.recurrence) {
      courseEvent.recurrence = event.recurrence[0]
      const dtStart = dateStringUTC(event.start.dateTime)
      const tmp = rrule.RRule.fromString(`DTSTART:${dtStart}\n${event.recurrence[0]}`)
      courseEvent.recurText = decodeRecurText(tmp)
      courseEvent.recurDates = decodeRecurDates(tmp)
    }
    // Maybe used later if we schedule in Calendar
    // const entryPoints = getNested(event, 'conferenceData', 'entryPoints')
    // if (entryPoints && entryPoints[0].entryPointType === 'video') {
    //   const el = entryPoints[0]
    //   courseEvent.isVideo = true
    //   courseEvent.meetingUri = el.uri
    //   courseEvent.meetingLabel = el.label
    //   courseEvent.meetingCode = el.meetingCode
    //   courseEvent.meetingPassword = el.password
    // }
    return courseEvent
  }

  const sortByTypeAndDate = (a, b) => {
    //    (a.type > b.type) ? 1 : (a.type === b.type) ? (new Date(a.start.dateTime || a.start.date).getTime() -
    //        new Date(b.start.dateTime || b.start.date).getTime()) : -1
    if (a.type === b.type) {
      return
      new Date(a.start.dateTime || a.start.date).getTime() -
        new Date(b.start.dateTime || b.start.date).getTime()
    }
    let comparison = 0
    if (a.type > b.type) {
      comparison = 1
    } else if (a.type < b.type) {
      comparison = -1
    }
    return comparison
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
  courseEvents.sort(sortByTypeAndDate)
  // courseEvents.forEach((e) => console.log(`${e.type} - ${e.id} - ${e.summary}`))
  //  console.log("\n\nException")
  //  exception.forEach(e => console.log(`${e.type} - ${e.id} - ${e.summary}`))
  //  console.log("\n\nRecurrent")
  //  recurrent.forEach(e => console.log(`${e.type} - ${e.id} - ${e.summary}`))

  return courseEvents
}

/**
 * search for existence of nested object key and return value if found
 * @see https://stackoverflow.com/questions/2631001/test-for-existence-of-nested-javascript-object-key
 * @param {object} obj object to search
 * @returns {object} value of the property else null
 */
function getNested(obj, ...args) {
  return args.reduce((obj, level) => obj && obj[level], obj)
}

const monthNames = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
]

const splitDate = (t = new Date()) => t.toLocaleString().split(/[^\d]/)

const fmtDate = (dtStr) => {
  if (typeof dtStr === 'undefined' || dtStr === '') return 'unknown'
  const [d, m, y] = splitDate(new Date(dtStr))
  return d + '-' + monthNames[m - 1] + '-' + y
}

const decodeRecurText = (rule) => rule.toText()

//get up to 10  dates for the recurring event
const decodeRecurDates = (eventRule) => {
  // console.log(`decodeDates: ${JSON.stringify(eventRule, null, 2)}`)
  // just return dd-mmm (max 5)
  const futureDates = eventRule
    .all((date, i) => i < 10)
    .map((dte) => fmtDate(new Date(dte)).slice(0, 6))
  return `${futureDates.join(', ')}${futureDates.length > 9 ? '...' : ''}`
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

const dateStringUTC = (dateTime) => {
  const date = new Date(dateTime)
  const mm = date.getUTCMonth() + 1
  return [
    date.getUTCFullYear().toString().padStart(4, '0'),
    mm.toString().padStart(2, '0'),
    date.getUTCDate().toString().padStart(2, '0'),
    'T',
    date.getUTCHours().toString().padStart(2, '0'),
    date.getUTCMinutes().toString().padStart(2, '0'),
    date.getUTCSeconds().toString().padStart(2, '0'),
    'Z',
  ].join('')
}
