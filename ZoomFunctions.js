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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zoomSheet = ss.getSheetByName("Calendar Download");

  const selectedRange = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  selectedRange.activate();
  const selection = zoomSheet.getSelection();
  const firstColumn = selection.getActiveRange().getColumn();
  const lastColumn = selection.getActiveRange().getLastColumn();

  // Must select one column only and must be column "C" (3)
  if (firstColumn != lastColumn || firstColumn != 3) {
    showToast(
      'You need to Select one/some Summary on the "Calendar Download" sheet',
      20
    );
    return;
  }
  const headers = zoomSheet
    .getRange(1, 1, 1, zoomSheet.getLastColumn())
    .getValues()
    .shift();
  const firstRow = selectedRange.getRow();
  const lastRow = selectedRange.getLastRow();
  const data = selectedRange
    .offset(0, -2, lastRow - firstRow + 1, headers.length)
    .getValues();

  const meetingOptions = {
    type: 2,
    duration: 30,
    timezone: "Australia/Sydney",
    password: "u3a",
    agenda: "Zoom session testing from Sheets",
    settings: {
      use_pmi: true,
      auto_recording: "none",
      mute_upon_entry: true,
      join_before_host: true,
    },
  };

  data.forEach((row) => {
    meetingOptions.topic = row[headers.indexOf("summary")];
    const start_time = row[headers.indexOf("startDateTime")];
    const end_time = row[headers.indexOf("endDateTime")];
    const duration = dataDiffMinutes(end_time, start_time);
    meetingOptions.duration = duration;

    s = start_time.toLocaleString().split(/[^\d]/);
    meetingOptions.start_time = `${s[2]}-${s[1]}-${s[0]}T${s[4]}:${s[5]}:${s[6]}`;

    const request = UrlFetchApp.fetch(
      `https://api.zoom.us/v2/users/${getZoomUserId()}/meetings`,
      {
        method: "POST",
        contentType: "application/json",
        headers: { Authorization: `Bearer ${getZoomAccessToken()}` },
        payload: JSON.stringify(meetingOptions),
      }
    );
    const { join_url, id } = JSON.parse(request.getContentText());
    Logger.log(`Zoom meeting ${id} created`, join_url);
  });

  function dataDiffMinutes(dte1, dte2) {
    const d1 = new Date(dte1);
    const d2 = new Date(dte2);
    let diff = (d2.getTime() - d1.getTime()) / 1000;
    return Math.abs(Math.round(diff / 60));
  }
}
