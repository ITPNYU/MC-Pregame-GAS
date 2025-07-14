const garage = 8; // Column H
const audioLab = 16; // Column P
const garageLightingNeeds = 31; // Column AE
const garageAudioNeeds = 32; // Column AF
const audioLabNeeds = 33; // Column AG
const recurring = 20; // Column T
const recurringOnMonday = 21; // Column U
const recurringOnTuesday = 22; // Column V
const recurringOnWednesday = 23; // Column W
const recurringOnThursday = 24; // Column X
const recurringOnFriday = 25; // Column Y
const timeEdit = [
  recurringOnMonday, // Column U
  recurringOnTuesday, // Column V
  recurringOnWednesday, // Column W
  recurringOnThursday, // Column X
  recurringOnFriday, // Column Y
];

// Get rules from DataValidationRules sheet
const DATA_VALIDATION_RULE_NAME = "DataValidationRules";
const audioLabRule = "A2";
const garageLightingRule = "B2";
const garageAudioRule = "C2";
const checkboxRule = "D2";

function onEdit(e) {
  const bookingInfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const sheet = bookingInfoSheet[1];
  const rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_VALIDATION_RULE_NAME);
  const editedCell = e.range;

  const row = editedCell.getRow();
  const newValue = editedCell.getValue();

  if (editedCell.getColumn() === garage) {
    handleEdit(row, newValue, garage, sheet, rulesSheet, "");
  } else if (editedCell.getColumn() === audioLab) {
    handleEdit(row, newValue, audioLab, sheet, rulesSheet, "");
  } else if (editedCell.getColumn() === audioLabNeeds) {
    deleteEdit(row, audioLab, audioLabNeeds, sheet, "Audio Lab Not Selected");
  } else if (editedCell.getColumn() === garageLightingNeeds) {
    deleteEdit(row, garage, garageLightingNeeds, sheet, "Garage Not Selected");
  } else if (editedCell.getColumn() === garageAudioNeeds) {
    deleteEdit(row, garage, garageAudioNeeds, sheet, "Garage Not Selected");
  } else if (editedCell.getColumn() === recurring) {
    handleEdit(row, newValue, recurring, sheet, rulesSheet, "");
  } else if (timeEdit.includes(editedCell.getColumn())) {
    deleteEdit(row, recurring, editedCell.getColumn(), sheet, "Recurring Weekly Not Selected");
  }
}

function deleteEdit(row, referenceColumn, column, sheet, message) {
  const referenceValue = sheet.getRange(row, referenceColumn).getValue();
  if (referenceValue == false) removeDropdowns([column], row, sheet, message);
}

function handleEdit(row, newValue, column, sheet, rulesSheet, message) {
  if (column === garage) {
    if (newValue === true) {
      const sourceRanges = [garageLightingRule, garageAudioRule];
      const restoreColumns = [garageLightingNeeds, garageAudioNeeds];
      restoreDropdowns(restoreColumns, sourceRanges, row, sheet, rulesSheet);
    } else if (newValue === false) {
      const targetColumns = [garageLightingNeeds, garageAudioNeeds];
      removeDropdowns(targetColumns, row, sheet, message);
    }
  } else if (column === audioLab) {
    if (newValue === true) {
      const sourceRanges = [audioLabRule];
      const restoreColumns = [audioLabNeeds];
      restoreDropdowns(restoreColumns, sourceRanges, row, sheet, rulesSheet);
    } else if (newValue === false) {
      const targetColumns = [audioLabNeeds];
      removeDropdowns(targetColumns, row, sheet, message);
    }
  } else if (column === recurring) {
    if (newValue == true) {
      const sourceRanges = [checkboxRule, checkboxRule, checkboxRule, checkboxRule, checkboxRule];
      const restoreColumns = [...timeEdit];
      restoreDropdowns(restoreColumns, sourceRanges, row, sheet, rulesSheet);
    } else if (newValue == false) {
      const targetColumns = [...timeEdit];
      removeDropdowns(targetColumns, row, sheet, message);
    }
  }
}

function removeDropdowns(targetColumns, row, sheet, message) {
  for (let i = 0; i < targetColumns.length; i++) {
    var targetRange = sheet.getRange(row, targetColumns[i]);
    // Remove data validation (disable editing) for the target cell,
    // Optionally clear the value
    targetRange.clearDataValidations();
    targetRange.setValue("");

    // Show error message with Alert
    if (message !== "") {
      const indexToColumn = (index) => {
        if (index > 26) {
          return String.fromCharCode(65) + String.fromCharCode(65 + index - 27);
        } else {
          return String.fromCharCode(65 + index - 1);
        }
      };
      SpreadsheetApp.getUi().alert(`Row ${row}, Column ${indexToColumn(targetColumns[i])}: ${message}`);
    }
  }
}

function restoreDropdowns(restoreColumns, sourceRanges, row, sheet, rulesSheet) {
  for (let j = 0; j < sourceRanges.length; j++) {
    var sourceRange = rulesSheet.getRange(sourceRanges[j]);
    var restoreRange = sheet.getRange(row, restoreColumns[j]);

    // Restore dropdown in target cell if condition is not met,
    // Optionally clear the value
    restoreRange.setValue("");

    const rule = sourceRange.getDataValidation();
    restoreRange.setDataValidation(rule); // Restore the dropdown
    restoreRange.setBackground("white"); // Reset background color
  }
}
