
function onEdit(event) {
  if (!event || (event.range.columnStart === 1 && event.range.rowStart === 1)) return;
  
  const sheet = event.source.getActiveSheet();
  const range = event.range;
  const row = range.getRow();
  const column = range.getColumn();
  
  if (sheet.getName() === "Dropdowns") {
    syncGpaSheet();
  } else if (column === 4) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Stop', 'You cannot edit IDs', ui.ButtonSet.OK);
    range.setValue(event.oldValue);
    return;
  } else if (column === 1 && sheet.getRange("A1").getValue() === "GPA" && row > 1) {
    if (sheet.getName() === "Eurorelief Database") {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Stop', 'You may only set GPAs on the Working List.', ui.ButtonSet.OK);
      range.setValue(event.oldValue);
      return;
    } else if (getGpaSheetNames().includes(sheet.getName()) || sheet.getName() === "Working List") {
      handleGpaEdit(event);
    }
    return;
  } else {
    handleSyncEdit(event);
  }
}

function handleSyncEdit(event) {
  const sheet = event.source.getActiveSheet();
  const range = event.range;
  const row = range.getRow();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Syncing your edits...', 'Hey there :)', 2);
  
  const id = sheet.getRange(row, 4).getValue();
  const databaseSheet = spreadsheet.getSheetByName("Eurorelief Database");
  const gpaSheetNames = getGpaSheetNames();
  const targetSheet = gpaSheetNames.includes(sheet.getName()) || sheet.getName() === "Working List" 
                      ? databaseSheet 
                      : spreadsheet.getSheetByName(sheet.getRange(row, 1).getValue());
  
  if (targetSheet) {
    const dataRange = targetSheet.getDataRange();
    const values = dataRange.getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][3] === id) {
        targetSheet.getRange(i + 1, range.getColumn()).setValue(range.getValue());
        break;
      }
    }
  }
}

function handleGpaEdit(event) {
  if (!event) return;
  const sheet = event.source.getActiveSheet();
  const range = event.range;
  const row = range.getRow();
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmation', 'Do you have the authorization to set this GPA?', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.NO) {
    range.setValue(event.oldValue);
    return;
  } else if (response == ui.Button.YES) {
    const newSheetName = range.getValue();
    const newSheet = newSheetName === "None"
      ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Working List")
      : SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
    const destinationRow = newSheet.getLastRow() + 1;
    const destinationRange = newSheet.getRange(destinationRow, 1, 1, newSheet.getLastColumn());
    
    handleSyncEdit(event);
    if (newSheetName !== "None" && destinationRange) {
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).copyTo(destinationRange, {formatOnly: false});
    }
    sheet.deleteRow(row);
  }
}

function getGpaSheetNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dropdowns");
  const lastRow = sheet.getLastRow();
  return sheet.getRange(3, 1, lastRow - 2, 1).getValues().flat();
}

function onRefresh() {
  removeFilters();
  updateWorkingList();
  syncGpaSheet();
}

function removeFilters() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => sheet.getFilter() && sheet.getFilter().remove());
}

function syncGpaSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Dropdowns");
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
  
  values.forEach(name => {
    if (name[0].length > 0 && !spreadsheet.getSheetByName(name)) {
      const newSheet = spreadsheet.insertSheet(name[0]);
      const ui = SpreadsheetApp.getUi();
      ui.alert('Heads up!', 'Creating a new sheet for ' + name[0], ui.ButtonSet.OK);
      const sourceSheet = spreadsheet.getSheetByName("Eurorelief Database");
      sourceSheet.getRange("A1:AD1").copyTo(newSheet.getRange("A1:AD1"));
      newSheet.hideColumns(4);
      newSheet.hideColumns(20, 7);
    }
  });
}

function updateWorkingList() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const databaseSheet = spreadsheet.getSheetByName("Eurorelief Database");
  const workingListSheet = spreadsheet.getSheetByName("Working List");
  workingListSheet.clear();
  
  databaseSheet.getRange("A2:AE").sort({column: 1, ascending: true});
  const values = databaseSheet.getRange("A2:AE" + databaseSheet.getLastRow()).getValues();
  
  const filteredRows = values.filter(row => row.some(cell => cell === "None"));
  
  if (filteredRows.length > 0) {
    workingListSheet.getRange(2, 1, filteredRows.length, values[0].length).setValues(filteredRows);
  }
  
  const sourceSheet = spreadsheet.getSheetByName("Eurorelief Database");
  sourceSheet.getRange("A1:AE1").copyTo(workingListSheet.getRange("A1:AE1"));
  workingListSheet.getRange(2, 1, workingListSheet.getLastRow() - 1, workingListSheet.getLastColumn())
    .sort([{column: 6, ascending: true}, {column: 5, ascending: true}]);
}

function deleteFromWorkingList() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const workingListSheet = spreadsheet.getSheetByName("Working List");
  const databaseSheet = spreadsheet.getSheetByName("Eurorelief Database");
  
  if (!workingListSheet || !databaseSheet) {
    Logger.log("Sheet not found.");
    return;
  }
  
  const workingListData = workingListSheet.getDataRange().getValues();
  const databaseData = databaseSheet.getDataRange().getValues();
  
  const workingListIdIndex = getColumnIndex(workingListSheet, "ID");
  const databaseIdIndex = getColumnIndex(databaseSheet, "ID");
  const gpaColumnIndex = getColumnIndex(databaseSheet, "GPA");
  
  if (workingListIdIndex === -1 || databaseIdIndex === -1 || gpaColumnIndex === -1) {
    Logger.log("Column not found.");
    return;
  }
  
  for (let i = workingListData.length - 1; i > 0; i--) {
    const idValue = workingListData[i][workingListIdIndex - 1];
    const databaseRow = databaseData.find(row => row[databaseIdIndex - 1] === idValue && row[gpaColumnIndex - 1] !== "");
    
    if (databaseRow) {
      workingListSheet.deleteRow(i + 1);
    }
  }
}

function getColumnIndex(sheet, header) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(header) + 1;
}

function removeNames() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const toRemoveSheet = spreadsheet.getSheetByName("To Remove");
  const toRemoveData = toRemoveSheet.getDataRange().getValues();
  const namesToRemove = toRemoveData.map(row => row[0]).filter(name => name);
  
  const databaseSheet = spreadsheet.getSheetByName("Eurorelief Database");
  const databaseData = databaseSheet.getDataRange().getValues();
  const pastVisitorsSheet = spreadsheet.getSheetByName("Past Visitors");
  
  for (let i = databaseData.length - 1; i >= 0; i--) {
    if (namesToRemove.includes(databaseData[i][30])) {
      pastVisitorsSheet.appendRow(databaseData[i]);
      databaseSheet.deleteRow(i + 1);
    }
  }
}

function overnightSync() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dropdownSheet = spreadsheet.getSheetByName("Dropdowns");
  const lastRow = dropdownSheet.getLastRow();
  const values = dropdownSheet.getRange(3, 1, lastRow - 2, 1).getValues();
  
  const databaseSheet = spreadsheet.getSheetByName("Eurorelief Database");
  const databaseIds = databaseSheet.getRange("D:D").getValues().flat();
  
  values.forEach(row => {
    const name = row[0];
    const targetSheet = spreadsheet.getSheetByName(name);
    if (targetSheet) {
      const data = targetSheet.getData

Range().getValues();
      data.forEach(row => {
        const id = row[3];
        if (id === "ID") return;
        const databaseRow = databaseIds.indexOf(id);
        if (databaseRow !== -1) {
          databaseSheet.getRange(databaseRow + 1, 1, 1, row.length).setValues([row]);
        }
      });
    }
  });
}
