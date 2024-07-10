var spreadsheetApp = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();

var databaseSheet = spreadsheetApp.getSheetByName("Database");
var parametersSheet = spreadsheetApp.getSheetByName("Parameters");
var historySheet = spreadsheetApp.getSheetByName("History");

var currentJobs = getValuesFromRange(parametersSheet, 2, 2);
var jobSheets = getValuesFromRange(parametersSheet, 2, 3);
var jobDictionary = buildDictionaryFromColumns("D", "C");

function onEdit(e) {
  handleEdit(e);
}

function handleEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  var column = range.getColumn();
  var value = range.getValue();

  var hiredColumn = getHeaderColumn(sheet, "Hired");
  if (hiredColumn && column == hiredColumn && row >= 2) {
    handleHiredColumnEdit(sheet, row, column, value);
  }

  var inCampColumn = getHeaderColumn(sheet, "In Camp");
  if (inCampColumn && jobSheets.includes(sheet.getName()) && column == inCampColumn && row >= 2) {
    handleInCampColumnEdit(sheet, row, column, value);
  }

  if (sheet.getName() == "Interested" && column == inCampColumn && row >= 2) {
    handleInterestedSheetEdit(sheet, row, range, value);
  }
}

function handleHiredColumnEdit(sheet, row, column, value) {
  var jobColumn = column + 1;
  if (value === true) {
    var response = ui.alert('Confirmation', 'Do you want to hire this resident?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      moveResidentToJobSheet(sheet, row, jobColumn);
    } else {
      range.setValue(false);
    }
  } else if (value === false && jobSheets.includes(sheet.getName())) {
    var response = ui.alert('Confirmation', 'Is this resident finished working?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      moveResidentToHistory(sheet, row);
    } else {
      range.setValue(true);
    }
  }
}

function handleInCampColumnEdit(sheet, row, column, value) {
  if (value === false) {
    var response = ui.alert('Confirmation', 'Did this resident leave camp?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      moveResidentToHistory(sheet, row);
    } else {
      range.setValue(true);
    }
  }
}

function handleInterestedSheetEdit(sheet, row, range, value) {
  if (value === false) {
    var response = ui.alert('Confirmation', 'Do you want to remove this entry?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      sheet.deleteRows(row);
    } else {
      range.setValue(true);
    }
  }
}

function moveResidentToJobSheet(sheet, row, jobColumn) {
  var jobValue = sheet.getRange(row, jobColumn).getValue();
  var destinationSheetName = jobDictionary[jobValue] || "Other";
  var destinationSheet = spreadsheetApp.getSheetByName(destinationSheetName);

  copyRowToSheet(sheet, row, destinationSheet);
  sheet.deleteRows(row);
}

function moveResidentToHistory(sheet, row) {
  copyRowToSheet(sheet, row, historySheet);
  sheet.deleteRows(row);
}

function copyRowToSheet(sourceSheet, row, targetSheet) {
  var lastRow = targetSheet.getLastRow();
  var nextRow = lastRow + 1;

  targetSheet.insertRowAfter(lastRow);
  var sourceRange = sourceSheet.getRange(row, 1, 1, sourceSheet.getLastColumn());
  var targetRange = targetSheet.getRange(nextRow, 1, 1, sourceSheet.getLastColumn());

  targetRange.clear();
  sourceRange.copyTo(targetRange);
  targetSheet.showRows(nextRow);
}

function getHeaderColumn(sheet, header) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnIndex = headers.indexOf(header);
  return columnIndex + 1;
}

function buildDictionaryFromColumns(keyColumn, valueColumn) {
  var keys = getValuesFromRange(parametersSheet, 1, keyColumn);
  var values = getValuesFromRange(parametersSheet, 1, valueColumn);
  return buildDictionary(keys, values);
}

function buildDictionary(keys, values) {
  var dictionary = {};
  for (var i = 0; i < keys.length; i++) {
    dictionary[keys[i]] = values[i];
  }
  return dictionary;
}

function getValuesFromRange(sheet, startRow, column) {
  return sheet.getRange(startRow, column, sheet.getLastRow() - startRow + 1, 1).getValues().flat();
}

function refreshPage() {
  sortInterestedSheet();
  sortMceWceSheet();
}

function sortInterestedSheet() {
  var sheet = spreadsheetApp.getActiveSheet();
  if (sheet.getName() === "Interested") {
    var jobColumnIndex = getHeaderColumn(sheet, "Job");
    var dateColumnIndex = getHeaderColumn(sheet, "Date Interested");

    if (jobColumnIndex && dateColumnIndex) {
      var sortingRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      sortingRange.sort([{ column: jobColumnIndex, ascending: false }, { column: dateColumnIndex, ascending: true }]);
    }
  }
}

function sortMceWceSheet() {
  var sheet = spreadsheetApp.getActiveSheet();
  if (sheet.getName() === "WCE/MCE") {
    var jobColumnIndex = getHeaderColumn(sheet, "Job");
    if (jobColumnIndex) {
      var sortingRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      sortingRange.sort([{ column: jobColumnIndex, ascending: false }]);
    }
  }
}
