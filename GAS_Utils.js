function getMyFolder(sheetObj) {
  return DriveApp.getFileById(sheetObj.getId()).getParents().next()
}

function checkIfFolderExistElseCreate(parent, folderName) {
  let folder
  try {
    folder = parent.getFoldersByName(folderName).next()
  } catch (e) {
    folder = parent.createFolder(folderName)
  }
  return folder
}

function showToast(msg, duration) {
  SpreadsheetApp.getActive().toast(msg, 'U3A Courses', duration)
}

// return an object with rangeNames and full A1 notation addresses ("named Range": 'Sheet'!A1:B2)
function getNamedRangesA1(ss) {
  let sheetIdToName = {}
  ss.getSheets().forEach(function (e) {
    sheetIdToName[e.getSheetId()] = e.getSheetName()
  })
  let result = {}
  res = Sheets.Spreadsheets.get(ss.getId(), {
    fields: 'namedRanges',
  }).namedRanges.forEach(function (e) {
    const id = e.range.sheetId || 0
    const sRow = e.range.startRowIndex || 0
    const eRow = e.range.endRowIndex || 1
    const sCol = e.range.startColumnIndex || 0
    const eCol = e.range.endColumnIndex || 1
    const sheetName = sheetIdToName[id.toString()]
    const a1notation = ss
      .getSheetByName(sheetName)
      .getRange(sRow + 1, sCol + 1, eRow - sRow, eCol - sCol)
      .getA1Notation()
    result[e.name] = `'${sheetName}'!${a1notation}`
  })
  return result
}

// Takes in
//   spreadsheet:(Range Object)
//   fileName:(String)
// returns
//   pdfFile:(File Object)
//
function makePDFfromRange(selectedRange, fileName, folderName) {
  const sheetToExport = selectedRange.getSheet()
  const ss = sheetToExport.getParent()
  const myFolder = getMyFolder(ss)
  const folder = checkIfFolderExistElseCreate(myFolder, folderName)

  const fileList = folder.getFilesByName(fileName)
  if (fileList.hasNext()) {
    fileList.next().setTrashed(true)
  }
  const url = ss.getUrl()
  const rangeParam =
    '&r1=' +
    (selectedRange.getRow() - 1) +
    '&r2=' +
    selectedRange.getLastRow() +
    '&c1=' +
    (selectedRange.getColumn() - 1) +
    '&c2=' +
    selectedRange.getLastColumn()

  const sheetParam = '&gid=' + selectedRange.getSheet().getSheetId()

  const exportUrl =
    url.replace(/\/edit.*$/, '') +
    '/export?exportFormat=pdf&format=pdf' +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +
    '&top_margin=0.75' +
    '&bottom_margin=0.75' +
    '&left_margin=0.7' +
    '&right_margin=0.7' +
    '&sheetnames=false&printtitle=false' +
    '&pagenum=false' +
    '&gridlines=false' +
    '&fzr=FALSE' +
    sheetParam +
    rangeParam

  //  Logger.log("exportUrl=" + exportUrl);
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  })

  const blob = response.getBlob()

  blob.setName(fileName)
  const pdfFile = folder.createFile(blob)

  showToast(`PDF Created as : ${pdfFile.getName()}`, 10)

  return pdfFile
}

/**
 * Create an object from a 2 dimensional array (usually sheet data)
 * from https://stackoverflow.com/questions/47555347/creating-a-json-object-from-google-sheets
 *
 * @param {array} data 2 dimensional array of rows with headings in first row
 * @returns {array} of objects with keys from row 1 with values from each other row
 *
 */
function getJsonArrayFromData(data) {
  const result = []
  const headers = data[0]
  const cols = headers.length

  for (var i = 1, l = data.length; i < l; i++) {
    // get a row to fill the object
    const row = data[i]
    // clear object
    const obj = {}
    for (var col = 0; col < cols; col++) {
      // fill object with new values
      obj[headers[col]] = row[col]
    }
    // add object in a final result
    result.push(obj)
  }

  return result
}

/**
 * Invokes a function, performing up to 5 retries with exponential backoff.
 * Retries with delays of approximately 1, 2, 4, 8 then 16 seconds for a total of
 * about 32 seconds before it gives up and rethrows the last error.
 * See: https://developers.google.com/google-apps/documents-list/#implementing_exponential_backoff
 * Author: peter.herrmann@gmail.com (Peter Herrmann)
 * Calls an anonymous function that concatenates a greeting with the current Apps user's email
 * var example1 = GASRetry.call(function(){return "Hello, " + Session.getActiveUser().getEmail();});
 * Calls an existing function
 * var example2 = GASRetry.call(myFunction);
 * Calls an anonymous function that calls an existing function with an argument
 * var example3 = GASRetry.call(function(){myFunction("something")});
 * Calls an anonymous function that invokes DocsList.setTrashed on myFile and logs retries with the Logger.log function.
 * var example4 = GASRetry.call(function(){myFile.setTrashed(true)}, Logger.log);
 *
 * @param {Function} func The anonymous or named function to call.
 * @param {Function} optLoggerFunction Optionally, you can pass a function that will be used to log to in the case of a retry.
 *                                     For example, Logger.log (no parentheses) will work.
 * @return {*} The value returned by the called function.
 */
function expBackOff(func, optLoggerFunction) {
  for (var n = 0; n < 6; n++) {
    try {
      return func()
    } catch (e) {
      if (optLoggerFunction) {
        optLoggerFunction('GASRetry ' + n + ': ' + e)
      }
      if (n == 5) {
        throw e
      }
      Utilities.sleep(Math.pow(2, n) * 1000 + Math.round(Math.random() * 1000))
    }
  }
}
