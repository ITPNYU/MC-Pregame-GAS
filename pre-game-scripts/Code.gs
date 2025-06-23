const DRAFT_CALENDAR_IDS = "RoomInfo";
const MAIN_CALENDAR_IDS = "MainCalendars";

function onOpen() {
  const ui = SpreadsheetApp.getUi(); // Get the user interface

  const email = Session.getActiveUser().getEmail();

  // if (email === "jg5626@nyu.edu") {
  //   ui.createMenu("Calendar") // Create a new menu
  //     .addItem("Push To Draft Calendars", "pushToDraftCalendars") // Add a menu item
  //     .addToUi(); // Add the menu to the UI

  //   ui.createMenu("Finalize")
  //     .addItem("Push To Main Calendars", "pushToMainCalendars") // Add item to the submenu
  //     .addToUi(); // Add the menu to the UI
  // }
  if (isDevTeam(email)) createDevTeamMenu(ui);
  else if (isAdmin(email)) createAdminMenu(ui);
}

function createDevTeamMenu(ui) {
  ui.createMenu("Calendar")
    .addItem("Write all Pregame Calendar classes", "writeClassesToPregame")
    .addItem("Write all Pregame Calendar events", "writeEventsToPregame")
    .addItem("Delete all Pregame Calendar classes", "deleteClassesFromPregame")
    .addItem("Delete all Pregame Calendar events", "deleteEventsFromPregame")
    .addItem("Delete all Draft Calendar classes", "deleteClassesFromDraft")
    .addItem("Delete all Draft Calendar events", "deleteEventsFromDraft")
    .addToUi();
}

function createAdminMenu(ui) {
  ui.createMenu("Calendar")
    .addItem("Write all Class rows to Draft Calendars", "writeClassesToDraft")
    .addItem("Write all Event rows to Draft Calendars", "writeEventsToDraft")
    .addItem("Delete all Draft Calendar classes", "deleteClassesFromDraft")
    .addItem("Delete all Draft Calendar events", "deleteEventsFromDraft")
    .addToUi();
}

function isDevTeam(email) {
  const devTeam = ["bs4396@nyu.edu", "nnp278@nyu.edu", "rh3555@nyu.edu"];
  return devTeam.includes(email);
}

function isAdmin(email) {
  const adminTeam = ["jg5626@nyu.edu"];
  return adminTeam.includes(email);
}

function pushToDraftCalendars() {
  pushToCalendar(true, DRAFT_CALENDAR_IDS);
}

function pushToMainCalendars() {
  pushToCalendar(false, MAIN_CALENDAR_IDS);
}

function pushToCalendar(eventsWillDelete, calendarInfoSheetName) {
  const roomInfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    calendarInfoSheetName
  );
  const bookingInfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  createCalendarEvents(roomInfoSheet, bookingInfoSheet[1], eventsWillDelete);
}

function createCalendarEvents(
  roomInfoSheet,
  bookingInfoSheet,
  eventsWillDelete
) {
  console.log(eventsWillDelete);

  var requesterEmails = getColumnValues(bookingInfoSheet, "B2:B");
  var eventTitles = getColumnValues(bookingInfoSheet, "G2:G");
  var eventRooms = getConcatenatedColumnData(bookingInfoSheet, "H", "S");
  var departments = getColumnValues(bookingInfoSheet, "D2:D");
  var descriptionInfo = getDescriptionInfo(bookingInfoSheet, "AD", "AM");
  var weeklyValues = getConcatenatedColumnData(bookingInfoSheet, "U", "Y");

  var startDateCells = getColumnValues(bookingInfoSheet, "Z2:Z");
  var startTimeCells = getColumnValues(bookingInfoSheet, "AA2:AA");
  var endDateCells = getColumnValues(bookingInfoSheet, "AB2:AB");
  var endTimeCells = getColumnValues(bookingInfoSheet, "AC2:AC");
  var isRecurring = getColumnValues(bookingInfoSheet, "T2:T");

  var roomNames = getRoomColumnValues(roomInfoSheet, "A:A");
  var roomCalendarIds = getRoomColumnValues(roomInfoSheet, "B:B");
  var roomNumbers = getRoomColumnValues(roomInfoSheet, "C:C");

  var roomNumberMap = createIDHashMap(roomNames, roomNumbers);
  const roomCalendarMap = createHashMap(roomNames, roomCalendarIds);
  const roomCalendarIDMap = createIDHashMap(roomNames, roomCalendarIds);

  var { startTimes, endTimes } = getStartAndEndTimes(
    startDateCells,
    startTimeCells,
    endTimeCells
  );

  if (eventsWillDelete == true) {
    for (let i = 0; i < roomCalendarIds.length; i++) {
      try {
        const roomCalendar = roomCalendarMap[i];
        deleteEvents(roomCalendar);
      } catch (error) {
        console.log(error);
      }
    }
  }

  for (let i = 0; i + 1 < bookingInfoSheet.getMaxRows(); i++) {
    i = skipEmptyRows(bookingInfoSheet, i);
    if (i + 2 > bookingInfoSheet.getMaxRows()) break;

    var recurrenceRule = null;
    if (isRecurring[i]) {
      recurrenceRule = getRecurrenceRule(endDateCells[i], weeklyValues[i]);

      const validWeekdays = getValidWeekdays(weeklyValues[i]);
      const { date: tempStartTime, offset } = getFirstValidWeekday(
        startTimes[i],
        validWeekdays
      );

      startTimes[i] = tempStartTime;

      var tempEndTime = new Date(endTimes[i]);

      tempEndTime.setHours(0, 0, 0, 0);

      tempEndTime.setDate(tempEndTime.getDate() + offset);

      if (tempEndTime <= endDateCells[i]) {
        endTimes[i].setDate(endTimes[i].getDate() + offset);
      }
    }

    let inviteeCalendars = [];
    let numberArray = [];
    var inviteRoom = null;

    for (let j = 0; j < eventRooms[i].length; j++) {
      const rooms = eventRooms[i];

      if (rooms[j] == true) {
        numberArray.push(roomNumberMap[j]);
        if (numberArray.length == 1) inviteRoom = j;
        else inviteeCalendars.push(roomCalendarIDMap[j]);
      }
    }

    const title =
      numberArray.join(", ") + "  " + departments[i] + "  " + eventTitles[i];
    const eventDescription =
      '<h2 style="font-size:24px; font-weight:bold;">Requester Details </h2>' +
      requesterEmails[i] +
      '<h2 style="font-size:24px; font-weight:bold;">Reservation Details </h2>' +
      descriptionInfo[i].map((item) => `• ${item}`).join("\n") +
      '<h2 style="font-size:24px; font-weight:bold;">Cancellation Policy </h2>' +
      "To cancel reservations please return to the Booking Tool, visit My Bookings, and click" +
      " 'cancel' on the booking at least 24 hours before the date of the event. Failure to cancel an unused booking is considered a no-show and may result in restricted use of the space.";

    const startDate = new Date(startDateCells[i]);
    const endDate = new Date(endDateCells[i]);
    var isError = false;

    try {
      if (recurrenceRule) {
        roomCalendarMap[inviteRoom].createEventSeries(
          title,
          startTimes[i],
          endTimes[i],
          recurrenceRule,
          {
            guests: inviteeCalendars.join(","), // Calendar emails as guests
            sendInvites: true,
            description: eventDescription,
          }
        );
      } else if (
        startDate.getTime() === endDate.getTime() ||
        endDateCells[i] == ""
      ) {
        roomCalendarMap[inviteRoom].createEvent(
          title,
          startTimes[i],
          endTimes[i],
          {
            guests: inviteeCalendars.join(","), // Calendar emails as guests
            sendInvites: true,
            description: eventDescription,
          }
        );
      } else if (startDate < endDate) {
        const { dailyStartTimes, dailyEndTimes, recurrenceRule } =
          getLongEventInfo(startDate, endDate, startTimes[i], endTimes[i]);

        if (recurrenceRule) {
          roomCalendarMap[inviteRoom].createEvent(
            title,
            dailyStartTimes[0],
            dailyEndTimes[0],
            {
              guests: inviteeCalendars.join(","), // Calendar emails as guests
              sendInvites: true,
              description: eventDescription,
            }
          );
          roomCalendarMap[inviteRoom].createEventSeries(
            title,
            dailyStartTimes[1],
            dailyEndTimes[1],
            recurrenceRule,
            {
              guests: inviteeCalendars.join(","), // Calendar emails as guests
              sendInvites: true,
              description: eventDescription,
            }
          );
          roomCalendarMap[inviteRoom].createEvent(
            title,
            dailyStartTimes[2],
            dailyEndTimes[2],
            {
              guests: inviteeCalendars.join(","), // Calendar emails as guests
              sendInvites: true,
              description: eventDescription,
            }
          );
        } else {
          roomCalendarMap[inviteRoom].createEvent(
            title,
            dailyStartTimes[0],
            dailyEndTimes[0],
            {
              guests: inviteeCalendars.join(","), // Calendar emails as guests
              sendInvites: true,
              description: eventDescription,
            }
          );
          roomCalendarMap[inviteRoom].createEvent(
            title,
            dailyStartTimes[1],
            dailyEndTimes[1],
            {
              guests: inviteeCalendars.join(","), // Calendar emails as guests
              sendInvites: true,
              description: eventDescription,
            }
          );
        }
      } else if (startDate > endDate) {
        isError = true;
        handleError(bookingInfoSheet, i);
      }

      if (!isError)
        bookingInfoSheet
          .getRange(i + 2, 1, 1, bookingInfoSheet.getLastColumn())
          .setBackground("white");
    } catch (error) {
      console.log(error);

      handleError(bookingInfoSheet, i);
    }
  }
}

function handleError(bookingInfoSheet, i) {
  bookingInfoSheet
    .getRange(i + 2, 1, 1, bookingInfoSheet.getLastColumn())
    .setBackground("red");
}

function getLongEventInfo(startDate, endDate, startTime, endTime) {
  let dailyStartTimes = [];
  let dailyEndTimes = [];

  let differenceInMs = endDate - startDate;

  var recurrenceRule = null;
  var offset = 0;
  // Convert milliseconds to days, hours, or other units as needed
  let differenceInDays = differenceInMs / (1000 * 60 * 60 * 24); // days

  let numberOfDays = differenceInDays + 1;

  let newStartTime = new Date(startDate); // Clone the start date for the end date
  newStartTime.setHours(0, 0, 0, 0); // Set to 00:00:00:00 AM

  let newEndTime = new Date(startDate); // Clone the start date for the end date
  newEndTime.setHours(23, 59, 59, 999); // Set to 11:59:59.999 PM

  dailyStartTimes.push(new Date(startTime));
  dailyEndTimes.push(new Date(newEndTime)); // Add to dailyEndTimes

  if (numberOfDays > 2) {
    newStartTime.setDate(newStartTime.getDate() + 1);
    newEndTime.setDate(newEndTime.getDate() + 1);
    offset = 1;
    dailyStartTimes.push(new Date(newStartTime)); // Add to the array
    dailyEndTimes.push(new Date(newEndTime));

    let recurrenceEndDate = new Date(endDate);
    recurrenceEndDate.setDate(recurrenceEndDate.getDate() - 1);

    recurrenceRule = CalendarApp.newRecurrence();

    recurrenceRule.addDailyRule().until(new Date(recurrenceEndDate));
  }

  let lastStartTime = new Date(newStartTime);
  lastStartTime.setDate(newStartTime.getDate() + differenceInDays - offset);

  dailyStartTimes.push(new Date(lastStartTime));
  endTime.setDate(endTime.getDate() + differenceInDays);
  dailyEndTimes.push(new Date(endTime)); // Add to the array

  return {
    dailyStartTimes: dailyStartTimes,
    dailyEndTimes: dailyEndTimes,
    recurrenceRule: recurrenceRule,
  };
}

// function getLongEventInfo(startDate, endDate, startTime, endTime) {

//   let dailyStartTimes = [];
//   let dailyEndTimes = [];

//   let differenceInMs = endDate - startDate;

//   // Convert milliseconds to days, hours, or other units as needed
//   let differenceInDays = differenceInMs / (1000 * 60 * 60 * 24);  // days

//   for (let j = 0; j <= differenceInDays; j++) {
//     let newStartTime = new Date(startDate); // Clone the start date for the end date
//     newStartTime.setHours(0, 0, 0, 0);  // Set to 00:00:00:00 AM
//     newStartTime.setDate(newStartTime.getDate() + j);

//     let newEndTime = new Date(startDate); // Clone the start date for the end date
//     newEndTime.setHours(23, 59, 59, 999);  // Set to 11:59:59.999 PM
//     newEndTime.setDate(newEndTime.getDate() + j);
//     if (j == 0) {
//       dailyStartTimes.push(startTime);

//       dailyEndTimes.push(newEndTime); // Add to dailyEndTimes
//     }
//     else if (j == differenceInDays) {
//       dailyStartTimes.push(newStartTime);

//       endTime.setDate(endTime.getDate() + j); // Add j days to the start date
//       dailyEndTimes.push(endTime); // Add to the array
//     }
//     else {
//       dailyStartTimes.push(newStartTime); // Add to the array

//       dailyEndTimes.push(newEndTime);
//     }
//   }

//   return { dailyStartTimes, dailyEndTimes };

// }

function getValidWeekdays(weeklyValues) {
  let validWeekdays = [];
  var daysOfWeek = true;

  for (let i = 0; i < weeklyValues.length; i++) {
    if (weeklyValues[i] == true) {
      daysOfWeek = false;

      validWeekdays.push(i);
    }
  }

  if (daysOfWeek == true) {
    return null;
  }

  return validWeekdays;
}

function getFirstValidWeekday(startDate, validWeekdays) {
  // Return the start date if no specific weekdays are set
  if (!validWeekdays || validWeekdays.length === 0)
    return { date: new Date(startDate), offset: 0 };

  let date = new Date(startDate);
  let offset = 0;

  // Loop until we find a matching weekday
  while (true) {
    // Adjust the day to match your specified system: 0 = Monday, ..., 6 = Sunday
    const adjustedDay = (date.getDay() + 6) % 7; // Adjusts: Sunday (0) becomes 6, Monday (1) becomes 0, ..., Saturday (6) becomes 5

    if (validWeekdays.includes(adjustedDay)) {
      // Return when we find the first valid weekday

      return { date, offset };
    }

    // Move to the next day if no match
    date.setDate(date.getDate() + 1);
    offset++;
  }
}

function skipEmptyRows(bookingInfoSheet, i) {
  const numRows = bookingInfoSheet.getLastRow() - (i + 2) + 1;

  var row = bookingInfoSheet.getRange(
    i + 2,
    1,
    numRows,
    bookingInfoSheet.getLastColumn()
  );

  var rowData = row.getValues();

  var backgrounds = row.getBackgrounds();

  var rowIsEmpty = checkRowData(rowData[0]);

  if (rowIsEmpty) {
    if (backgrounds[0][0].toLowerCase() === "#ff0000") {
      bookingInfoSheet
        .getRange(i + 2, 1, 1, bookingInfoSheet.getLastColumn())
        .setBackground("white");
    }
  }

  var index = 1;

  while (rowIsEmpty) {
    i++;
    if (i + 2 > bookingInfoSheet.getLastRow()) break;
    if (backgrounds[index][0].toLowerCase() === "#ff0000") {
      bookingInfoSheet
        .getRange(i + 2, 1, 1, bookingInfoSheet.getLastColumn())
        .setBackground("white");
    }

    rowIsEmpty = checkRowData(rowData[index]);
    index++;
  }

  return i;
}

function checkRowData(rowData) {
  var isEmpty = true;
  for (let j = 0; j < rowData.length; j++) {
    if (
      rowData[j] != "" &&
      rowData[j] != false &&
      !String(rowData[j]).includes("N/A")
    ) {
      isEmpty = false;
    }
  }

  return isEmpty;
}

function deleteEvents(roomCalendar) {
  const allEvents = roomCalendar.getEvents(
    new Date("2000-01-01"),
    new Date("2100-01-01")
  );

  const uniqueEvents = new Map();

  for (let i = allEvents.length - 1; i >= 0; i--) {
    const event = allEvents[i];
    const eventKey = `${event.getTitle()}-${event.getStartTime()}`;

    try {
      if (!uniqueEvents.has(eventKey)) {
        // Add to unique map and delete it
        uniqueEvents.set(eventKey, true);

        if (event.isRecurringEvent()) {
          // Delete the entire recurring series
          event.getEventSeries().deleteEventSeries();
        } else {
          // Delete single non-recurring events
          event.deleteEvent();
        }
      } else {
        // If it's a duplicate event, try deleting it directly
        event.deleteEvent();
      }
    } catch (e) {}
  }
}

// Helper function to generate the recurrence rule with a specific end date
function getRecurrenceRule(tempEndDate, weeklyValues) {
  const endDate = new Date(tempEndDate);
  endDate.setDate(endDate.getDate() + 1); //inclusive end date

  try {
    return daysOfWeekRecurrence(endDate, weeklyValues);
  } catch (error) {
    console.log(error);
  }
}

function daysOfWeekRecurrence(endDate, weeklyValues) {
  const recurrence = CalendarApp.newRecurrence();

  var daysOfWeek = true;

  for (let i = 0; i < weeklyValues.length; i++) {
    if (weeklyValues[i] == true) {
      daysOfWeek = false;

      recurrence
        .addWeeklyRule()
        .onlyOnWeekday(weekIndexToWeekday(i))
        .until(endDate);
    }
  }

  if (daysOfWeek == true) {
    recurrence.addWeeklyRule().until(endDate);
  }

  return recurrence;
}

function weekIndexToWeekday(i) {
  switch (i) {
    case 0:
      return CalendarApp.Weekday.MONDAY;
    case 1:
      return CalendarApp.Weekday.TUESDAY;
    case 2:
      return CalendarApp.Weekday.WEDNESDAY;
    case 3:
      return CalendarApp.Weekday.THURSDAY;
    case 4:
      return CalendarApp.Weekday.FRIDAY;
  }
}

function getStartAndEndTimes(startDateCells, startTimeCells, endTimeCells) {
  // Initialize arrays to store merged start and end times
  let startTimes = [];
  let endTimes = [];

  // Loop through all the dates and times, combining them into Date objects
  for (let i = 0; i < startDateCells.length; i++) {
    if (startTimeCells[i] == "") {
      startTimes.push("");
      endTimes.push("");
      continue;
    }
    // Create start time by merging the start date with the start time
    let startTime = new Date(startDateCells[i]);
    startTime.setHours(
      startTimeCells[i].getHours(),
      startTimeCells[i].getMinutes()
    );

    // Create end time by merging the end date with the end time
    let endTime = new Date(startDateCells[i]);
    endTime.setHours(endTimeCells[i].getHours(), endTimeCells[i].getMinutes());

    // Store the merged Date objects in the arrays
    startTimes.push(startTime);
    endTimes.push(endTime);
  }

  return { startTimes, endTimes };
}

function createHashMap(names, ids) {
  // Create the HashMap (using a JavaScript object)
  let hashMap = {};
  for (let i = 0; i < names.length; i++) {
    hashMap[names[i]] = CalendarApp.getCalendarById(ids[i]);
  }

  return hashMap;
}

function createIDHashMap(names, ids) {
  // Create the HashMap (using a JavaScript object)
  let hashMap = {};
  for (let i = 0; i < names.length; i++) {
    hashMap[names[i]] = ids[i];
  }

  return hashMap;
}

function getColumnValues(sheet, column) {
  // Get all values in column A (you can change the range if needed)
  const columnValues = sheet.getRange(column).getValues().flat();

  return columnValues;
}

function getRoomColumnValues(sheet, column) {
  // Get all values in column A (you can change the range if needed)
  const columnValues = sheet.getRange(column).getValues();

  // Filter out any empty rows (if needed) and flatten to 1D array
  const valuesArray = columnValues
    .flat() // Converts 2D array to 1D
    .filter((value) => value !== ""); // Removes empty cells

  return valuesArray;
}

function getDescriptionInfo(sheet, startColumn, endColumn) {
  // Convert column letters to numeric indices
  const startColumnIndex = columnLetterToIndex(startColumn);
  const endColumnIndex = columnLetterToIndex(endColumn);

  // Calculate the number of columns in the range
  const numberOfColumns = endColumnIndex - startColumnIndex + 1;

  // Get the range of values from the specified columns and rows
  const dataRange = sheet
    .getRange(1, startColumnIndex, sheet.getMaxRows(), numberOfColumns)
    .getValues();

  // Prepare the result: 2D array of concatenated strings
  const concatenatedData = [];

  // Loop through each row (starting from the 2nd row since 1st is title)
  for (let row = 1; row < sheet.getMaxRows(); row++) {
    const rowArray = [];

    // Loop through each column in the range
    for (let col = 0; col < numberOfColumns; col++) {
      const title = dataRange[0][col]; // First row = title
      const value = dataRange[row][col];
      if (value == "") rowArray.push(`${title}: ${false}`);
      // Format as "Title: Value" // Current row value
      else rowArray.push(`${title}: ${value}`); // Format as "Title: Value"
    }

    concatenatedData.push(rowArray); // Add the row's array to the result
  }

  return concatenatedData;
}

function getConcatenatedColumnData(sheet, startColumn, endColumn) {
  // Convert column letters to numeric indices
  const startColumnIndex = columnLetterToIndex(startColumn);
  const endColumnIndex = columnLetterToIndex(endColumn);

  // Calculate the number of columns in the range
  const numberOfColumns = endColumnIndex - startColumnIndex + 1;

  // Get the range of values from the specified columns and rows
  const dataRange = sheet
    .getRange(1, startColumnIndex, sheet.getMaxRows(), numberOfColumns)
    .getValues();

  // Prepare the result: 2D array of concatenated strings
  const concatenatedData = [];

  // Loop through each row (starting from the 2nd row since 1st is title)
  for (let row = 1; row < sheet.getMaxRows(); row++) {
    const rowArray = [];

    // Loop through each column in the range
    for (let col = 0; col < numberOfColumns; col++) {
      const value = dataRange[row][col]; // Current row value
      rowArray.push(value); // Format as "Title: Value"
    }

    concatenatedData.push(rowArray); // Add the row's array to the result
  }

  return concatenatedData;
}

// Helper function: Convert column letter to index (A = 1, B = 2, etc.)
function columnLetterToIndex(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + (letter.charCodeAt(i) - "A".charCodeAt(0) + 1);
  }
  return column;
}
