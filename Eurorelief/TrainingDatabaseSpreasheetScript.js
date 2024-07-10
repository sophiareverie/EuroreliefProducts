function onOpen() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    const importSheet = ss.getSheetByName("Importing");
    const databaseSheet = ss.getSheetByName("History");
    const schedulingSheet = ss.getSheetByName("Scheduling");
    const excludedSheet = ss.getSheetByName("People Excluded");
    const overviewSheet = ss.getSheetByName("Overview");
    const trainingsSheet = ss.getSheetByName("Edit Trainings");
    
    const importedData = getImportedData(importSheet);
    const databaseNames = getColumnValues(databaseSheet, "B:B");
    const scheduleNames = getColumnValues(schedulingSheet, "B:B");
    const excludedNames = getColumnValues(excludedSheet, "A2:A");
    
    const newDatabaseData = filterNewData(importedData, databaseNames, excludedNames);
    const newSchedulingData = filterNewData(importedData, scheduleNames, excludedNames, true);
    
    appendDataToSheet(databaseSheet, newDatabaseData);
    appendDataToSheet(schedulingSheet, newSchedulingData);
    
    removeOldAndDuplicateEntries(schedulingSheet);
    sortSheetByColumn(schedulingSheet, 4);
    
    manageTrainings(overviewSheet, trainingsSheet, schedulingSheet, ui);
    
    sortSheetByColumn(trainingsSheet, 1);
  }
  
  function getImportedData(sheet) {
    const rangeA = sheet.getRange("A:A").getValues();
    const rangeDE = sheet.getRange("D:E").getValues();
    const rangeN = sheet.getRange("N:N").getValues();
    const rangeP = sheet.getRange("P:P").getValues();
    
    return rangeA.map((row, index) => row.concat(rangeDE[index], rangeN[index], rangeP[index]));
  }
  
  function getColumnValues(sheet, range) {
    return sheet.getRange(range).getValues().flat();
  }
  
  function filterNewData(importedData, existingNames, excludedNames, isScheduling = false) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const newData = [];
    
    for (let i = 1; i < importedData.length; i++) {
      const [id, name, , , date] = importedData[i];
      if (!existingNames.includes(name) && !excludedNames.includes(name) && id && id !== "Visitor" && name && name !== "i58 Volunteer" && (!isScheduling || date > today)) {
        newData.push(importedData[i]);
        existingNames.push(name);
      }
    }
    return newData;
  }
  
  function appendDataToSheet(sheet, data) {
    if (data.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    }
  }
  
  function removeOldAndDuplicateEntries(sheet) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const schedulingData = sheet.getDataRange().getValues();
    const seenNames = {};
    
    for (let i = schedulingData.length - 1; i >= 1; i--) {
      const [ , name, , , date] = schedulingData[i];
    
      if (date instanceof Date && date < today) {
        sheet.deleteRow(i + 1);
        continue;
      }
    
      if (seenNames[name]) {
        sheet.deleteRow(i + 1);
      } else {
        seenNames[name] = true;
      }
    }
  }
  
  function sortSheetByColumn(sheet, column) {
    const lastRow = sheet.getLastRow();
    const sortRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    sortRange.sort(column);
  }
  
  function manageTrainings(overviewSheet, trainingsSheet, schedulingSheet, ui) {
    const overviewData = overviewSheet.getDataRange().getValues();
    const overviewTrainings = overviewData[0];
    const currentTrainings = getColumnValues(trainingsSheet, "A2:A");
    
    removeOldTrainings(overviewSheet, schedulingSheet, overviewTrainings, currentTrainings, ui);
    addNewTrainings(overviewSheet, schedulingSheet, overviewTrainings, currentTrainings, ui);
  }
  
  function removeOldTrainings(overviewSheet, schedulingSheet, overviewTrainings, currentTrainings, ui) {
    for (let i = overviewTrainings.length - 1; i >= 0; i--) {
      const training = overviewTrainings[i];
      if (training && !currentTrainings.includes(training)) {
        const response = ui.alert('Just checking...', `Do you want to remove the training ${training}?`, ui.ButtonSet.YES_NO);
        if (response == ui.Button.YES) {
          try {
            overviewSheet.deleteColumn(i + 1);
            schedulingSheet.deleteColumns(schedulingSheet.getRange("1:1").getValues()[0].indexOf(training), 3);
          } catch (e) {
            Logger.log(`Error deleting column: ${i + 1}`);
          }
        }
      }
    }
  }
  
  function addNewTrainings(overviewSheet, schedulingSheet, overviewTrainings, currentTrainings, ui) {
    let insertIndex = overviewTrainings.length;
    let lastSchedCol = schedulingSheet.getLastColumn();
    for (let i = 0; i < currentTrainings.length; i++) {
      const training = currentTrainings[i];
      if (!training) break;
      if (!overviewTrainings.includes(training)) {
        const response = ui.alert('Just checking...', `Do you want to add the training ${training}?`, ui.ButtonSet.YES_NO);
        if (response == ui.Button.YES) {
          try {
            schedulingSheet.insertColumnsAfter(lastSchedCol, 3);
            schedulingSheet.getRange(1, lastSchedCol + 1).setValue(training);
            schedulingSheet.getRange(2, lastSchedCol + 1, schedulingSheet.getLastRow(), 1).clearDataValidations().clearContent();
            schedulingSheet.getRange(1, lastSchedCol + 2).setValue("Scheduled");
            schedulingSheet.getRange(1, lastSchedCol + 3).setValue("Completed");
            lastSchedCol += 3;
            overviewSheet.insertColumnAfter(insertIndex);
            insertIndex++;
            overviewSheet.getRange(1, insertIndex).setValue(training);
            const colLetter = columnNumberToLetter(schedulingSheet.getRange("1:1").getValues()[0].indexOf(training) + 3);
            overviewSheet.getRange(2, insertIndex).setFormula(`=FILTER(Scheduling!B:B, {Scheduling!${colLetter}:${colLetter} = FALSE} * {Scheduling!D:D-TODAY()<7}*{Scheduling!E:E>TODAY()})`);
          } catch (e) {
            Logger.log(`Error inserting column: ${e}`);
          }
        }
      }
    }
  }
  
  function columnNumberToLetter(columnNumber) {
    let result = '';
    while (columnNumber > 0) {
      columnNumber--;
      result = String.fromCharCode(columnNumber % 26 + 'A'.charCodeAt(0)) + result;
      columnNumber = Math.floor(columnNumber / 26);
    }
    return result;
  }
  
  function onEdit(e) {
    const editedSheet = e.source.getActiveSheet();
    const editedRange = e.range;
    const editedRow = editedRange.getRow();
    const editedCol = editedRange.getColumn();
    const editedValue = editedRange.getValue();
    const columnHeader = editedSheet.getRange(1, editedCol).getValue();
  
    if (editedSheet.getName() === "Scheduling" && editedCol > 6 && columnHeader === "Completed") {
      handleSchedulingEdit(editedSheet, editedRow, editedCol, editedValue);
    }
  
    if (columnHeader === "SAVE:" && editedRow === 2 && editedCol === 2 && editedValue === true) {
      handleSaveEdit(editedSheet);
    }
  }
  
  function handleSchedulingEdit(editedSheet, editedRow, editedCol, editedValue) {
    const checkboxValue = editedValue === "TRUE";
    const trainingName = editedSheet.getRange(1, editedCol - 2).getValue();
    const name = editedSheet.getRange(editedRow, 2).getValue();
    const databaseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("History");
    const databaseData = databaseSheet.getDataRange().getValues();
  
    updateTrainingCompletion(databaseData, databaseSheet, name, trainingName, checkboxValue);
  }
  
  function updateTrainingCompletion(databaseData, databaseSheet, name, trainingName, addTraining) {
    for (let i = 0; i < databaseData.length; i++) {
      if (databaseData[i][1] === name) {
        const oldValue = databaseData[i][5];
        let newValue = "";
  
        if (addTraining) {
          if (!oldValue.toString().includes(trainingName)) {
            newValue = oldValue ? `${oldValue}, (${getTodaysDate()}) ${trainingName}` : `(${getTodaysDate()}) ${trainingName}`;
          } else {
            newValue = oldValue;
          }
        } else {
          newValue = oldValue.split(", ").filter(value => !value.includes(trainingName)).join(", ");
        }
  
        databaseSheet.getRange(i + 1, 6).setValue(newValue);
        break;
      }
    }
  }
  
  function handleSaveEdit(editedSheet) {
    const trainingName = editedSheet.getRange("B4").getValue();
    const action = editedSheet.getRange("B6").getValue();
    let dateToSave = editedSheet.getRange("B8").getValue();
  
    if (!trainingName || !dateToSave) {
      SpreadsheetApp.getUi().alert("Must have valid training name and date before saving.");
      editedSheet.getRange("B2").setValue(false);
      return;
    }
  
    dateToSave = Utilities.formatDate(dateToSave, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      `Are you sure you want to ${action.toLowerCase()} people for ${trainingName}?`,
      ui.ButtonSet.YES_NO
    );
  
    if (response == ui.Button.YES) {
      schedulePeopleForTraining(editedSheet, trainingName, dateToSave, action);
    }
  
    editedSheet.getRange("B2").setValue(false);
  }
  
  function schedulePeopleForTraining(editedSheet, trainingName, dateToSave, action) {
    const schedulingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scheduling");
    const headers = schedulingSheet.getRange(1, 1, 1, schedulingSheet.getLastColumn()).getValues()[0];
    const columnToFill = headers.indexOf(trainingName) + 1;
  
    if (columnToFill === 0) {
      Logger.log("Training name not found in headers");
      return;
    }
  
    const names = editedSheet.getRange("B10:B" + editedSheet.getLastRow()).getValues().flat();
    const schedulingNames = schedulingSheet.getRange("B:B").getValues().flat();
  
    names.forEach(name => {
      if (name) {
        const rowToUpdate = schedulingNames.indexOf(name) + 1;
        if (rowToUpdate > 0) {
          if (action === "Mark Scheduled") {
            schedulingSheet.getRange(rowToUpdate, columnToFill).setValue(dateToSave);
            schedulingSheet.getRange(rowToUpdate, columnToFill + 1).setValue(true);
          } else if (action === "Mark Completed") {
            schedulingSheet.getRange(rowToUpdate, columnToFill).setValue(dateToSave);
            schedulingSheet.getRange(rowToUpdate, columnToFill + 2).setValue(true);
          }
        }
      }
    });
  
    clearSaveForm(editedSheet);
  }
  
  function clearSaveForm(sheet) {
    sheet.getRange("B4").clearContent();
    sheet.getRange("B6").clearContent();
    sheet.getRange("B8").clearContent();
    sheet.getRange("B11:B").clearContent();
  }
  
  function getTodaysDate() {
    const today = new Date();
    return Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  
  function sync() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName("History");
    const schedSheet = ss.getSheetByName("Scheduling");
  
    const schedNames = getColumnValues(schedSheet, 2);
    const schedHeaders = schedSheet.getRange(1, 1, 1, schedSheet.getLastColumn()).getValues()[0];
    const historyData = historySheet.getDataRange().getValues();
  
    updateSchedulingSheet(historyData, schedNames, schedHeaders, schedSheet);
  }
  
  function updateSchedulingSheet(historyData, schedNames, schedHeaders, schedSheet) {
    for (let i = 0; i < historyData.length; i++) {
      const name = historyData[i][1];
      if (name) {
        const schedRowToUpdate = schedNames.indexOf(name) + 2; // Accounting for header row
        const trainings = historyData[i][5];
  
        if (trainings.length > 0) {
          updateTrainings(trainings, schedRowToUpdate, schedHeaders, schedSheet);
        }
      }
    }
  }
  
  function updateTrainings(trainings, schedRowToUpdate, schedHeaders, schedSheet) {
    const trainingsList = trainings.split(", ");
  
    trainingsList.forEach(trainingInfo => {
      const date = trainingInfo.substring(1, 11);
      const training = trainingInfo.substring(13);
  
      const schedColToUpdate = schedHeaders.indexOf(training) + 1;
      if (schedRowToUpdate > 1 && schedColToUpdate > 0) {
        schedSheet.getRange(schedRowToUpdate, schedColToUpdate).setValue(date);
        schedSheet.getRange(schedRowToUpdate, schedColToUpdate + 2).setValue(true);
      }
    });
  }
  