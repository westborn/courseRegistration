const ZOOM_API_KEY = 'YNzuV58YTw23l1blpB05_Q'
const ZOOM_API_SECRET = 'udVIOcqFC3OBPWn9pbTgjVlrHUt3mYIyqNgf'
const ZOOM_EMAIL = 'u3acomputerclub@hotmail.com'

const getZoomAccessToken = () => {
  const encode = (text) => Utilities.base64Encode(text).replace(/=+$/, '')
  const header = { alg: 'HS256', typ: 'JWT' }
  const encodedHeader = encode(JSON.stringify(header))
  const payload = {
    iss: ZOOM_API_KEY,
    exp: Date.now() + 3600,
  }
  const encodedPayload = encode(JSON.stringify(payload))
  const toSign = `${encodedHeader}.${encodedPayload}`
  const signature = encode(Utilities.computeHmacSha256Signature(toSign, ZOOM_API_SECRET))
  return `${toSign}.${signature}`
}

const getZoomUserId = () => {
  const request = UrlFetchApp.fetch('https://api.zoom.us/v2/users/', {
    method: 'GET',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${getZoomAccessToken()}` },
  })
  const { users } = JSON.parse(request.getContentText())
  const [{ id } = {}] = users.filter(({ email }) => email === ZOOM_EMAIL)
  return id
}

function selectedZoomSessions() {
  const res = metaSelected('Calendar Download', 1)
  if (!res) {
    return
  }
  const { sheetSelected, rangeSelected, numRowsSelected } = res
  const headers = sheetSelected.getRange(1, 1, 1, sheetSelected.getLastColumn()).getValues().shift()
  const data = sheetSelected
    .getRange(rangeSelected)
    .offset(0, 0, numRowsSelected, headers.length)
    .getValues()

  const meetingOptions = {
    type: 2,
    duration: 30,
    timezone: 'Australia/Sydney',
    password: 'u3a',
    agenda: 'Zoom session testing from Sheets',
    settings: {
      use_pmi: true,
      auto_recording: 'none',
      mute_upon_entry: true,
      join_before_host: true,
    },
  }

  data.forEach((row) => {
    meetingOptions.topic = row[headers.indexOf('summary')]
    const start_time = row[headers.indexOf('startDateTime')]
    const end_time = row[headers.indexOf('endDateTime')]
    const duration = dateDiffMinutes(end_time, start_time)
    meetingOptions.duration = duration

    s = start_time
      .toLocaleString('en-AU', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false,
      })
      .split(/[^\d]/)
    meetingOptions.start_time = `${s[2]}-${s[1]}-${s[0]}T${s[4]}:${s[5]}:${s[6]}`
    const request = UrlFetchApp.fetch(`https://api.zoom.us/v2/users/${getZoomUserId()}/meetings`, {
      method: 'POST',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${getZoomAccessToken()}` },
      payload: JSON.stringify(meetingOptions),
    })
    // const { join_url, id } = JSON.parse(request.getContentText())
    // Logger.log(`Zoom meeting ${id} created`, join_url)
  })
}
