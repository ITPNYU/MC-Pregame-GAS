const garage = 8; // Column H
const audioLab = 16;

const audioLabNeeds = 31;
const garageLightingNeeds = 32;
const garageAudioNeeds = 33;

const recurring = 20;

let timeEdit = [];
timeEdit.push(21, 22, 23, 24, 25);

function onEdit(e) {
  const bookingInfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  const sheet = bookingInfoSheet[1];

  const roomInfoSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RoomInfo");

  const editedCell = e.range;

  // Define columns: Monitor column 7 (G), target column 17 (Q)

  const row = editedCell.getRow();
  const newValue = editedCell.getValue();

  // Check if the edited cell is in column G
  if (editedCell.getColumn() === garage) {
    handleEdit(row, newValue, garage, sheet, roomInfoSheet, "");
  } else if (editedCell.getColumn() === audioLab) {
    handleEdit(row, newValue, audioLab, sheet, roomInfoSheet, "");
  } else if (editedCell.getColumn() === audioLabNeeds) {
    deleteEdit(row, audioLab, audioLabNeeds, sheet, "Audio Lab Not Selected");
  } else if (editedCell.getColumn() === garageLightingNeeds) {
    deleteEdit(row, garage, garageLightingNeeds, sheet, "Garage Not Selected");
  } else if (editedCell.getColumn() === garageAudioNeeds) {
    deleteEdit(row, garage, garageAudioNeeds, sheet, "Garage Not Selected");
  } else if (editedCell.getColumn() === recurring) {
    handleEdit(row, newValue, recurring, sheet, roomInfoSheet, "");
  } else if (timeEdit.includes(editedCell.getColumn())) {
    deleteEdit(
      row,
      recurring,
      editedCell.getColumn(),
      sheet,
      "Recurring Weekly Not Selected"
    );
  }
}

function deleteEdit(row, referenceColumn, column, sheet, message) {
  let targetColumns = [];

  const referenceValue = sheet.getRange(row, referenceColumn).getValue();

  handleDelete(targetColumns, row, sheet, column, referenceValue, message);
}

function handleDelete(
  targetColumns,
  row,
  sheet,
  column,
  referenceValue,
  message
) {
  targetColumns.push(column);
  if (referenceValue == false) {
    removeDropdowns(targetColumns, row, sheet, message);
  }
}

function handleEdit(row, newValue, column, sheet, roomInfoSheet, message) {
  let restoreColumns = [];
  let targetColumns = []; // Column Q
  let sourceRanges = []; // Column Q

  if (column === garage) {
    if (newValue === true) {
      sourceRanges.push("E2", "F2");
      restoreColumns.push(32, 33); //should always be equal in length to sourceRanges
      restoreDropdowns(restoreColumns, sourceRanges, row, sheet, roomInfoSheet);
    } else if (newValue === false) {
      targetColumns.push(32, 33);
      removeDropdowns(targetColumns, row, sheet, message);
    }
  } else if (column === audioLab) {
    if (newValue === true) {
      sourceRanges.push("D2");
      restoreColumns.push(31); //should always be equal in length to sourceRanges
      restoreDropdowns(restoreColumns, sourceRanges, row, sheet, roomInfoSheet);
    } else if (newValue === false) {
      targetColumns.push(31);
      removeDropdowns(targetColumns, row, sheet, message);
    }
  } else if (column === recurring) {
    if (newValue == true) {
      sourceRanges.push("G2", "G2", "G2", "G2", "G2");
      restoreColumns.push(21, 22, 23, 24, 25); //should always be equal in length to sourceRanges
      restoreDropdowns(restoreColumns, sourceRanges, row, sheet, roomInfoSheet);
    } else if (newValue == false) {
      targetColumns.push(21, 22, 23, 24, 25);
      removeDropdowns(targetColumns, row, sheet, message);
    }
  }
}

// function removeAndRestoreDropdowns(restoreColumns, targetColumns, sourceRanges, row, sheet, roomInfoSheet, message) {

//   for (let i = 0; i < targetColumns.length; i++) {

//     var targetRange = sheet.getRange(row, targetColumns[i]);
//     removeDropdown(targetRange, message);

//   }

//   for (let j = 0; j < sourceRanges.length; j++) {

//     var sourceRange = roomInfoSheet.getRange(sourceRanges[j]);
//     var restoreRange = sheet.getRange(row, restoreColumns[j]);
//     restoreDropdown(restoreRange, sourceRange);

//   }

// }

function removeDropdowns(targetColumns, row, sheet, message) {
  for (let i = 0; i < targetColumns.length; i++) {
    var targetRange = sheet.getRange(row, targetColumns[i]);

    // Remove data validation (disable editing) for the target cell
    targetRange.clearDataValidations(); // Clear existing validations
    //targetRange.setBackground("#D9D9D9"); // Set background to grey
    if (message === "") {
      targetRange.setValue(""); // Optionally clear the value
    } else {
      targetRange.setValue("N/A" + " - " + message); // Optionally clear the value
    }
    //removeDropdown(targetRange, message);
  }
}

function restoreDropdowns(
  restoreColumns,
  sourceRanges,
  row,
  sheet,
  roomInfoSheet
) {
  for (let j = 0; j < sourceRanges.length; j++) {
    var sourceRange = roomInfoSheet.getRange(sourceRanges[j]);
    var restoreRange = sheet.getRange(row, restoreColumns[j]);

    // Restore dropdown in target cell if condition is not met
    restoreRange.setValue(""); // Optionally clear the value

    const rule = sourceRange.getDataValidation();

    restoreRange.setDataValidation(rule); // Restore the dropdown
    restoreRange.setBackground("white"); // Reset background color
  }
}
