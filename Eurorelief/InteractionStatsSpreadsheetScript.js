
function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    if (range.getA1Notation() === "E1" && range.getValue() === true) {
        handleSaveButton(sheet, range, e);
    }

    if (sheet.getName() === "Edit Departments") {
        handleDepartmentEdits(sheet);
    }
}

function handleSaveButton(sheet, range, event) {
    var ui = SpreadsheetApp.getUi();
    var weekString = sheet.getRange("A1").getValue().trim();
    
    if (!weekString) {
        ui.alert('Error', 'You must enter the dates of the week you are saving in cell A1.', ui.ButtonSet.OK);
        range.setValue(false);
        return;
    }
    
    var response = ui.alert('Confirmation', 'Are you sure you want to save and clear this week?', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
        saveAndClearWeek(sheet, event, weekString);
    }
    
    range.setValue(false);
}

function saveAndClearWeek(sheet, event, weekString) {
    var archiveSheet = event.source.getSheetByName("Archive");
    var dataRange = sheet.getRange("B2:B" + sheet.getLastRow());
    var values = dataRange.getValues();
    
    if (values.length > 0) {
        var lastColumn = archiveSheet.getLastColumn();
        archiveSheet.getRange(3, lastColumn + 1, values.length, 1).setValues(values);
    }
    
    var noteValues = sheet.getRange("C2:C" + sheet.getLastRow()).getValues().flat().filter(value => value !== "");
    var commaSeparatedString = noteValues.join(", ");
    archiveSheet.getRange(2, lastColumn + 1).setValue(commaSeparatedString);
    sheet.getRange("B2:C" + sheet.getLastRow()).clear();
    archiveSheet.getRange(1, lastColumn + 1).setValue(weekString);
}

function handleDepartmentEdits(sheet) {
    var departments = sheet.getRange("A2:A").getValues().flat().filter(value => value);
    var weeklySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly");
    var currentDepartments = weeklySheet.getRange("A2:A").getValues().flat();
    var archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Archive");
    var archiveDepartments = archiveSheet.getRange("A3:A").getValues().flat();

    addNewDepartments(departments, currentDepartments, weeklySheet, archiveDepartments, archiveSheet);
    removeDeletedDepartments(departments, currentDepartments, weeklySheet);

    sortSheets(weeklySheet, sheet, archiveSheet);
}

function addNewDepartments(departments, currentDepartments, weeklySheet, archiveDepartments, archiveSheet) {
    departments.forEach(dept => {
        if (!currentDepartments.includes(dept)) {
            weeklySheet.insertRowAfter(weeklySheet.getLastRow());
            weeklySheet.getRange("A" + (weeklySheet.getLastRow())).setValue(dept);
        }
        if (!archiveDepartments.includes(dept)) {
            archiveSheet.insertRowAfter(archiveSheet.getLastRow());
            archiveSheet.getRange("A" + (archiveSheet.getLastRow())).setValue(dept);
        }
    });
}

function removeDeletedDepartments(departments, currentDepartments, weeklySheet) {
    for (var i = currentDepartments.length - 1; i >= 0; i--) {
        var dept = currentDepartments[i];
        if (!departments.includes(dept)) {
            weeklySheet.deleteRow(i + 2);
        }
    }
}

function sortSheets(weeklySheet, editSheet, archiveSheet) {
    weeklySheet.getRange("A2:D").sort(1);
    editSheet.getRange("A2:A").sort(1);
    archiveSheet.getRange(3, 1, archiveSheet.getLastRow() - 2, archiveSheet.getLastColumn()).sort(1);
}
