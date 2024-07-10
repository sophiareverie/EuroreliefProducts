function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var saveColumnIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf("Save") + 1;
    var clearCheckboxCell = sheet.getRange(2, saveColumnIndex).getA1Notation();

    if (range.getA1Notation() === clearCheckboxCell && range.getValue() === true) {
        if (confirmAction('Are you sure you want to save and clear this timesheet?')) {
            alertUser('Please wait until the checkboxes are cleared, then change the day of the week in the dropdown.');
            saveAndClearTimesheet();
        }
        range.setValue(false);
    }
}

function confirmAction(message) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Confirmation', message, ui.ButtonSet.YES_NO);
    return response == ui.Button.YES;
}

function alertUser(message) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Wait', message, ui.ButtonSet.OK);
}

function saveAndClearTimesheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var timeSheet = spreadsheet.getSheetByName("TimeSheet");
    timeSheet.deleteColumn(6);
    insertEmptyColumn(timeSheet);
    insertDate(timeSheet);
    setAttendanceValue(timeSheet);
    copyFormulaContents(timeSheet);
    clearCheckboxes();
}

function insertEmptyColumn(sheet) {
    var lastColumnIndex = sheet.getLastColumn();
    sheet.insertColumnAfter(lastColumnIndex);
}

function insertDate(sheet) {
    var lastColumnIndex = sheet.getLastColumn();
    var lastColumnDate = sheet.getRange(1, lastColumnIndex - 1).getValue();
    var nextDay = new Date(lastColumnDate.getTime() + 24 * 60 * 60 * 1000);
    var dateHeader = Utilities.formatDate(nextDay, Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.getRange(1, lastColumnIndex).setValue(dateHeader);
}

function setAttendanceValue(sheet) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var attendanceSheet = spreadsheet.getSheetByName("Attendance");
    var attendanceValue = attendanceSheet.getRange("D1").getValue();
    var lastColumnIndex = sheet.getLastColumn();
    sheet.getRange(2, lastColumnIndex).setValue(attendanceValue);
}

function copyFormulaContents(sheet) {
    var lastColumnIndex = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var formulaRange = sheet.getRange(3, lastColumnIndex, lastRow - 2, 1);
    var formula = '=IFERROR(' +
        'IF(INDEX(Attendance!$A$1:$A$100, MATCH(A3, Attendance!$D$1:$D$100, 0)), "X", "")' +
        '& IF(INDEX(Attendance!$B$1:$B$100, MATCH(A3, Attendance!$D$1:$D$100, 0)), "LATE", "")' +
        '& IF(INDEX(Attendance!$F$1:$F$100, MATCH(A3, Attendance!$D$1:$D$100, 0)), "OK", "")' +
        '& IF(AND(NOT(INDEX(Attendance!$A$1:$A$100, MATCH(A3, Attendance!$D$1:$D$100, 0))), INDEX(Attendance!$C$1:$C$100, MATCH(A3, Attendance!$D$1:$D$100, 0))), "NO", "")' +
        ', "")';

    formulaRange.setFormula(formula);

    for (var i = 3; i <= lastRow; i++) {
        var cell = sheet.getRange(i, lastColumnIndex);
        var value = cell.getValue();
        cell.setValue(value);
    }
}

function clearCheckboxes() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var checkboxesRange = sheet.getRange(3, 1, sheet.getLastRow() - 2, 3);
    var checkboxesColumnF = sheet.getRange(3, 6, sheet.getLastRow() - 2, 1);
    checkboxesRange.setValue(false);
    checkboxesColumnF.setValue(false);
}
