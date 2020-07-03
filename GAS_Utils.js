function getMyFolder(sheetObj) {
  return DriveApp.getFileById(sheetObj.getId()).getParents().next();
}

function showToast(msg, duration) {
  SpreadsheetApp.getActive().toast(msg, "U3A Courses", duration);
}

// return an object with rangeNames and full A1 notation addresses ("named Range": 'Sheet'!A1:B2)
function getNamedRangesA1(ss) {
  var sheetIdToName = {};
  ss.getSheets().forEach(function (e) {
    sheetIdToName[e.getSheetId()] = e.getSheetName();
  });
  var result = {};
  res = Sheets.Spreadsheets.get(ss.getId(), {
    fields: "namedRanges",
  }).namedRanges.forEach(function (e) {
    var id = e.range.sheetId || 0;
    var sRow = e.range.startRowIndex || 0;
    var eRow = e.range.endRowIndex || 1;
    var sCol = e.range.startColumnIndex || 0;
    var eCol = e.range.endColumnIndex || 1;
    var sheetName = sheetIdToName[id.toString()];
    var a1notation = ss
      .getSheetByName(sheetName)
      .getRange(sRow + 1, sCol + 1, eRow - sRow, eCol - sCol)
      .getA1Notation();
    result[e.name] = `'${sheetName}'!${a1notation}`;
  });
  return result;
}

// Takes in
//   spreadsheet:(Range Object)
//   fileName:(String)
// returns
//   pdfFile:(File Object)
//
function makePDFfromRange(selectedRange, fileName) {
  var sheetToExport = selectedRange.getSheet();
  var ss = sheetToExport.getParent();
  var folder = getMyFolder(ss);

  var fileList = folder.getFilesByName(fileName);
  if (fileList.hasNext()) {
    fileList.next().setTrashed(true);
  }
  var url = ss.getUrl();
  var rangeParam =
    "&r1=" +
    (selectedRange.getRow() - 1) +
    "&r2=" +
    selectedRange.getLastRow() +
    "&c1=" +
    (selectedRange.getColumn() - 1) +
    "&c2=" +
    selectedRange.getLastColumn();

  var sheetParam = "&gid=" + selectedRange.getSheet().getSheetId();

  var exportUrl =
    url.replace(/\/edit.*$/, "") +
    "/export?exportFormat=pdf&format=pdf" +
    "&size=A4" +
    "&portrait=true" +
    "&fitw=true" +
    "&top_margin=0.75" +
    "&bottom_margin=0.75" +
    "&left_margin=0.7" +
    "&right_margin=0.7" +
    "&sheetnames=false&printtitle=false" +
    "&pagenum=false" +
    "&gridlines=false" +
    "&fzr=FALSE" +
    sheetParam +
    rangeParam;

  //  Logger.log("exportUrl=" + exportUrl);
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken(),
    },
  });

  var blob = response.getBlob();

  blob = blob.setName(fileName);
  var pdfFile = folder.createFile(blob);

  showToast(`PDF Created as : ${pdfFile.getName()}`, 10);

  return pdfFile;
}
