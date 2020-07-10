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

const createZoomMeeting = () => {
  //https://marketplace.zoom.us/docs/api-reference/zoom-api/meetings/meetingcreate#request-body
  const meetingOptions = {
    topic: 'Zoom Meeting created with Google Script',
    type: 2,
    start_time: '2020-07-30T10:45:00',
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

  const request = UrlFetchApp.fetch(`https://api.zoom.us/v2/users/${getZoomUserId()}/meetings`, {
    method: 'POST',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${getZoomAccessToken()}` },
    payload: JSON.stringify(meetingOptions),
  })
  const { join_url, id } = JSON.parse(request.getContentText())
  Logger.log(`Zoom meeting ${id} created`, join_url)
}
