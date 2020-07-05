function getCalendarList() {
  const calendars;
  let pageToken;
  const result = [];
  do {
    calendars = Calendar.CalendarList.list({
      maxResults: 100,
      pageToken: pageToken,
    });
    if (calendars.items && calendars.items.length > 0) {
      for (let i = 0; i < calendars.items.length; i++) {
        const calendar = calendars.items[i];
        result.push({ summary: calendar.summary, id: calendar.id });
        //        console.log('%s (ID: %s)', calendar.summary, calendar.id)
      }
    } else {
      console.log("No calendars found.");
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);
  return result;
}

// =============================================================================
// get all the events for the term from the calendar

function downloadCalendarEvents({term,  calendarId} = {}) {

  eval(
    UrlFetchApp.fetch(
      "https://jakubroztocil.github.io/rrule/dist/es5/rrule-tz.min.js"
    ).getContentText()
  );

  //  console.log('transcribeCalendar')
  //get details of where this app is running from
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOptions = ss.getSheetByName("Options");
  const sheetDownload = ss.getSheetByName("Calendar Download");

  // setup dates from the term selected (start and end dates)
  const terms = sheetOptions.getRange(1, 1, 5, 3).getValues();
  const eventRequest = {
      singleEvents: false,
      timeMin: new Date(terms[term][1]).toISOString(),
      timeMax: new Date(terms[term][2]).toISOString(),
    }

    //  get the events and format them
  const courseEvents =  retrieveCalendarEvents(calendarId, eventRequest)
  //  console.log(JSON.stringify(events, null, 2))

  //clear the sheet we are going to download the events to
  sheetDownload.insertRowBefore(2);
  const lastRow = sheetDownload.getLastRow();
  if (lastRow > 2) {
    sheetDownload.deleteRows(3, lastRow - 2);
  }

  if (!courseEvents.length) {
    sheetDownload.getRange(2, 1).setValue("No events Found");
    return;
  }
  //  console.log(courseEvents)
  const rows = courseEvents.map((d) => flatten_(d));
  //  console.log(rows)
  const heads = sheetDownload.getDataRange().offset(0, 0, 1).getValues()[0];

  // convert object data into a 2d array
  const tr = rows.map((row) => heads.map((key) => row[String(key)] || ""));

  // write result
  sheetDownload
    .getRange(sheetDownload.getLastRow() + 1, 1, tr.length, tr[0].length)
    .setValues(tr);

  return;
}

// ====================================================================================
//Get all necessary course events (dates, location, summary, description) as an array of objects
function retrieveCalendarEvents(calendarId, eventRequest) {

  const unpackEvent = (type, event) => {
    const courseEvent = {
      type: type,
      id: event.id || '',
      summary: event.summary || '',
      description: event.description.replace(/(<([^>]+)>)/gi, '') || '',
      location: event.location || '',
      startDateTime: new Date(event.start.dateTime).toLocaleString().replace(',', ''),
      endDateTime: new Date(event.end.dateTime).toLocaleString().replace(',', ''),
      recurrence: '',
      recurText: '',
      recurDates: '',
      presenter: getNested(event, 'extendedProperties', 'private', 'presenter'),
      contact: getNested(event, 'extendedProperties', 'private', 'contact'),
      min: getNested(event, 'extendedProperties', 'private', 'min'),
      max: getNested(event, 'extendedProperties', 'private', 'max'),
      cost: getNested(event, 'conferenceData', 'entryPoints', 'entryPointType'),
      isVideo: false,
      meetingUri: '',
      meetingLabel: '',
      meetingCode: '',
      meetingPassword: '',
    }

    if (event.recurrence) {
      courseEvent.recurrence = event.recurrence[0]
      const tmp = rrule.RRule.fromString(event.recurrence[0])
      courseEvent.recurText = decodeRecurText(tmp)
      courseEvent.recurDates = decodeRecurDates(tmp, '2020-06-01T12:00:00-10:00')
    }

    const entryPoints = getNested(event, 'conferenceData', 'entryPoints')
    if (entryPoints && entryPoints[0].entryPointType === 'video') {
      const el = entryPoints[0]
      courseEvent.isVideo = true
      courseEvent.meetingUri = el.uri
      courseEvent.meetingLabel = el.label
      courseEvent.meetingCode = el.meetingCode
      courseEvent.meetingPassword = el.password
    }
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
        courseEvents.push(unpackEvent('exception', event))
      } else courseEvents.push(unpackEvent('standalone', event))
    } else {
      courseEvents.push(unpackEvent('recurrent', event))
    }
  })

  const recurTypes = courseEvents.filter((event) => event.type === 'recurrent')
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
      const exists = courseEvents.find((obj) => obj.type === 'exception' && obj.id === el.id)
      if (!exists) {
        courseEvents.push(unpackEvent('instance', el))
      }
    })
  })
  console.log(courseEvents.length)
  courseEvents.sort(sortByTypeAndDate)
  courseEvents.forEach((e) => console.log(`${e.type} - ${e.id} - ${e.summary}`))
  //  console.log("\n\nException")
  //  exception.forEach(e => console.log(`${e.type} - ${e.id} - ${e.summary}`))
  //  console.log("\n\nRecurrent")
  //  recurrent.forEach(e => console.log(`${e.type} - ${e.id} - ${e.summary}`))

  return courseEvents
}


function getNested(obj, ...args) {
  return args.reduce((obj, level) => obj && obj[level], obj)
}
const monthNames = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

const splitDate = (t = new Date()) => t.toLocaleString().split(/[^\d]/);

const fmtDate = (dtStr) => {
  if (typeof dtStr === "undefined" || dtStr === "") return "unknown";
  const [d, m, y] = splitDate(new Date(dtStr));
  return d + "-" + monthNames[m - 1] + "-" + y;
};

const decodeRecurText = (rule) => rule.toText();

const decodeRecurDates = (eventRule, dte = new Date()) => {
  // console.log(`decodeDates: ${JSON.stringify(eventRule, null, 2)}`)
  const newRule = new rrule.RRule({
    ...eventRule.origOptions,
    dtstart: new Date(dte),
  });
  // just return dd-mmm (max 5)
  const futureDates = newRule
    .all((date, i) => i < 6)
    .map((dte) => fmtDate(new Date(dte)).slice(0, 6));
  return `${futureDates.join(", ")}${futureDates.length > 5 ? "..." : ""}`;
};
