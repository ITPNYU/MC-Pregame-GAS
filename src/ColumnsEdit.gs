const garage = 11; // Column K
const audioLab = 21; // Column U
const garageLightingNeeds = 40; // Column AN
const garageAudioNeeds = 41; // Column AO
const audioLabNeeds = 42; // Column AP
const recurring = 28; // Column AB
const recurringOnMonday = 29; // Column AC
const recurringOnTuesday = 30; // Column AD
const recurringOnWednesday = 31; // Column AE
const recurringOnThursday = 32; // Column AF
const recurringOnFriday = 33; // Column AG
const timeEdit = [
  recurringOnMonday, // Column AC
  recurringOnTuesday, // Column AD
  recurringOnWednesday, // Column AE
  recurringOnThursday, // Column AF
  recurringOnFriday, // Column AG
];

// Get rules from DataValidationRules sheet
const DATA_VALIDATION_RULE_NAME = "DataValidationRules";
const audioLabRule = "A2";
const garageLightingRule = "B2";
const garageAudioRule = "C2";
const checkboxRule = "D2";

function onEdit(e) {

  // ____Reccuring Events and Room Availability _______________________
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

  // _____Cascading Dropdown for Reservation Category and Type____
  const edited_sheet = e.source.getActiveSheet();
  const range = e.range;
  const sheetName = edited_sheet.getName();
  const MAIN_SHEET = sheet.getName();
  const DATA_SHEET = "PregameCategoryAndType";
  const SETTINGS_SHEET = "PregameSetting";



  // ── Handle Settings change ───────────────────────────────────
  // When Settings!A2 is edited, refresh every Type dropdown on Main
  if (sheetName === SETTINGS_SHEET && range.getRow() === 2 && range.getColumn() === 1) {
    refreshAllTypeDropdowns(e.source, MAIN_SHEET, DATA_SHEET);
    return;
  }

  // ── Handle Main Column H edit ────────────────────────────────
  if (sheetName !== MAIN_SHEET) return;
  // Column H is the 8th column
  if (range.getColumn() !== 8) return;
  // Skip the header row
  if (range.getRow() <= 1) return;


  const selectedCategory = e.value;
  // Column I is the 9th column
  var targetCell = edited_sheet.getRange(range.getRow(), 9);

  // Clear previous value & validation whenever the category changes
  targetCell.clearContent();
  targetCell.clearDataValidations();

  // User deleted the value, do nothing
  if (!selectedCategory) return;

  // ── 1. Read settings mode ────────────────────────────────────
  const mode = getSettingsMode(e.source);

  // ── 2. Build the hash map ────────────────────────────────────
  const dataSheet = e.source.getSheetByName(DATA_SHEET);

  const categoryMap = buildCategoryMap(dataSheet);

  // ── 3. Look up types for the selected category + mode ────────
  const lookupKey = mode + "|" + selectedCategory;
  const options = categoryMap[lookupKey] || [];

  // ── 4. Apply dropdown ────────────────────────────────────────
  if (options.length > 0) {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .build();
    targetCell.setDataValidation(rule);
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


/**
 * Builds a hash map from the Data sheet.
 * Key:   "{Pregame}|{Category}"  (e.g. "Semester|Class")
 * Value: Array of Type strings   (e.g. ["Exam", "Instruction"])
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @returns {Object.<string, string[]>}
 */
function buildCategoryMap(dataSheet) {
  var rows = dataSheet.getRange("A2:C" + dataSheet.getLastRow()).getValues();
  var map = {};

  for (var i = 0; i < rows.length; i++) {
    var pregame = String(rows[i][0]).trim();
    var category = String(rows[i][1]).trim();
    var type = String(rows[i][2]).trim();

    if (!pregame || !category || !type) continue; // skip blank rows

    var key = pregame + "|" + category;

    if (!map[key]) {
      map[key] = [];
    }
    map[key].push(type);
  }

  return map;
}

/**
 * Reads the Settings sheet and returns the current mode
 * ("Semester" or "Annual").
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {string}
 */
function getSettingsMode(ss) {
  var settingsSheet = ss.getSheetByName("PregameSetting");
  // The mode is stored in cell A2
  return String(settingsSheet.getRange("A2").getValue()).trim();
}

/**
 * Re-applies Type dropdowns on every populated row of Main
 * after a Settings change.
 *
 * For each row where Column A has a Category:
 *   - Rebuild the Column B dropdown using the NEW mode.
 *   - If the current Type value in Column B is no longer valid
 *     under the new mode, clear it.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} mainSheetName
 * @param {string} dataSheetName
 */
function refreshAllTypeDropdowns(ss, mainSheetName, dataSheetName) {
  var mainSheet = ss.getSheetByName(mainSheetName);
  var dataSheet = ss.getSheetByName(dataSheetName);
  var lastRow = mainSheet.getLastRow();

  if (lastRow <= 1) return; // nothing below the header

  var mode = getSettingsMode(ss);
  var categoryMap = buildCategoryMap(dataSheet);

  // Read columns E & F in one batch (rows 2 → lastRow)
  var categories = mainSheet.getRange(2, 8, lastRow - 1, 1).getValues();
  var types = mainSheet.getRange(2, 9, lastRow - 1, 1).getValues();

  for (var i = 0; i < categories.length; i++) {
    var category = String(categories[i][0]).trim();
    var row = i + 2; // 1-indexed, skip header
    var typeCell = mainSheet.getRange(row, 9);

    if (!category) {
      // No category → clear any leftover validation/value
      typeCell.clearContent();
      typeCell.clearDataValidations();
      continue;
    }

    var lookupKey = mode + "|" + category;
    var options = categoryMap[lookupKey] || [];

    if (options.length > 0) {
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options, true)
        .build();
      typeCell.setDataValidation(rule);

      // Clear the current value if it's not valid under the new mode
      var currentType = String(types[i][0]).trim();
      if (currentType && options.indexOf(currentType) === -1) {
        typeCell.clearContent();
      }
    } else {
      // No valid options for this category under the new mode
      typeCell.clearContent();
      typeCell.clearDataValidations();
    }
  }
}