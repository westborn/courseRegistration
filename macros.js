function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E13').activate();
  var sourceData = spreadsheet.getRange('B12:C473');
  var pivotTable = spreadsheet.getRange('E12').createPivotTable(sourceData);
  var pivotValue = pivotTable.addPivotValue(2, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('numberCourses');
  var pivotGroup = pivotTable.addRowGroup(2);
  sourceData = spreadsheet.getRange('B12:C473');
  pivotTable = spreadsheet.getRange('E12').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(2, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('numberCourses');
  pivotGroup = pivotTable.addRowGroup(2);
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['', 'David Monro', 'George Stone'])
  .build();
  pivotTable.addFilter(2, criteria);
  sourceData = spreadsheet.getRange('B12:C473');
  pivotTable = spreadsheet.getRange('E12').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(2, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('numberCourses');
  pivotGroup = pivotTable.addRowGroup(2);
  criteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['David Monro', 'George Stone'])
  .build();
  pivotTable.addFilter(2, criteria);
};

function UntitledMacro2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H12').activate();
  var sourceData = spreadsheet.getRange('B12:C79');
  var pivotTable = spreadsheet.getRange('H12').createPivotTable(sourceData);
  var pivotValue = pivotTable.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('numberAttendees');
  var pivotGroup = pivotTable.addRowGroup(3);
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['', 'Apple Q\'s on the Fly', 'Australia and the last Ice Age', 'Australia\'s Vietnam War', 'Book Chat', 'Critical Thinking', 'Demystifying Apple Technology', 'How to buy wine you like ONLINE', 'Jesus is Coming - Look Busy!', 'Mindfulness Meditation', 'Poetry Workshop', 'Storywriting', 'The Cold War', 'Why are we Post Fact?', 'Zoom I.T. Class'])
  .build();
  pivotTable.addFilter(3, criteria);
  sourceData = spreadsheet.getRange('B12:C79');
  pivotTable = spreadsheet.getRange('H12').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('numberAttendees');
  pivotGroup = pivotTable.addRowGroup(3);
  criteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['Apple Q\'s on the Fly', 'Australia and the last Ice Age', 'Australia\'s Vietnam War', 'Book Chat', 'Critical Thinking', 'Demystifying Apple Technology', 'How to buy wine you like ONLINE', 'Jesus is Coming - Look Busy!', 'Mindfulness Meditation', 'Poetry Workshop', 'Storywriting', 'The Cold War', 'Why are we Post Fact?', 'Zoom I.T. Class'])
  .build();
  pivotTable.addFilter(3, criteria);
  sourceData = spreadsheet.getRange('B12:C79');
  pivotTable = spreadsheet.getRange('H12').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('numberAttendees');
  pivotGroup = pivotTable.addRowGroup(3);
  sourceData = spreadsheet.getRange('B12:C78');
  pivotTable = spreadsheet.getRange('H12').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue.setDisplayName('numberAttendees');
  pivotGroup = pivotTable.addRowGroup(3);
};