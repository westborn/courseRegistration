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
