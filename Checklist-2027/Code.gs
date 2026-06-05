function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Checklist')
    .addItem('Doer Setup', 'doersetup')
    .addToUi();
}

function doersetup() {
  Browser.msgBox('Setup is completed');
}

// ===== Helpers (as-is) =====
Date.prototype.addTheDays = function(days) {
  var date = new Date(this.valueOf());
  date.setDate(date.getDate() + days);
  return date;
};

function getLastRowSpecial(range){
  var rowNum = 0;
  var blank = false;
  for (var row = 0; row < range.length; row++) {
    if (range[row][0] === "" && !blank) {
      rowNum = row;
      blank = true;
    } else if (range[row][0] !== "") {
      blank = false;
    }
  }
  return rowNum;
}

function findinB(name,data) {
  var valB = name;
  for (var nn = 0; nn < data.length; ++nn) {
    if (data[nn][1] == valB) { break; } // match in col B then break
  }
  return data[nn][0]; // show column A
}

function Return_Date(F_Date,num,hr,min){ //Return_Date(F_Date,num)
  var dt = new Date(F_Date.getTime() + 24*60*60*1000*(num+1));
  dt.setHours(hr);
  dt.setMinutes(min);
  return dt;
}

// ===========================

function createChecklist() {
  const tz = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Task List
  const taskListSheet = ss.getSheetByName("Task List");
  if (!taskListSheet) throw new Error("Sheet 'Task List' not found.");
  const taskLastRow = taskListSheet.getLastRow();
  if (taskLastRow < 2) return; // no data
  const taskListData = taskListSheet.getRange(2, 1, taskLastRow - 1, 5).getValues();

  // Setup Sheet
  let skipSundays = ss.getSheetByName("Setup Sheet").getRange("B32").getValue();
  if (!skipSundays) skipSundays = "Yes";

  // Master
  const masterSheet = ss.getSheetByName("Master");
  if (!masterSheet) throw new Error("Sheet 'Master' not found.");
  const masterTaskIDs = masterSheet.getRange("B:B").getValues(); // 2D
  const masterLastRow = masterTaskIDs.filter(r => r[0] !== "" && r[0] != null).length;
  let lastMasterTaskID = masterLastRow ? Number(masterTaskIDs[masterLastRow - 1][0]) : 0;

  const masterArray = [];

  // Working Day Calendar
  const calendarSheet = ss.getSheetByName("Working Day Calender");
  if (!calendarSheet) throw new Error("Sheet 'Working Day Calender' not found.");

  const calendarColA = calendarSheet.getRange("A:A").getValues(); // 2D
  const calendarLen = calendarColA.filter(r => r[0]).length;
  if (calendarLen < 2) return; // no data rows

  const allCalendarDates = calendarSheet.getRange(2, 1, calendarLen - 1, 1).getValues().flat();
  const calendarLastDate = calendarColA[calendarLen - 1][0];
  const workingDatesStr = allCalendarDates.map(x => Utilities.formatDate(new Date(x), tz, "yyyy-MM-dd"));

  // Iterate tasks
  taskListData.forEach(function(task, index) {
    let theTask = task[0];
    let theDoer = task[1];
    let theFreq = task[2]; // D,W,M,Q,H,Y,F
    let theDate = task[3];
    let theStatus = task[4];

    if (theStatus === "Sent") return;

    // ---- WEEKLY (W) ----
    if (theFreq === "W") {
      while (Date.compare(calendarLastDate, theDate) === 1) {
        lastMasterTaskID += 1;
        let endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        let frozenDate = theDate.clone();
        while (!workingDatesStr.includes(endDateStr)) {
          theDate = Date.parse(theDate).addDays(-1);
          endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        }
        masterArray.push([theDoer, lastMasterTaskID, theFreq, theTask, new Date(theDate)]);
        theDate = Date.parse(frozenDate).addWeeks(1);
      }

    // ---- MONTHLY (M) ----
    } else if (theFreq === "M") {
      while (Date.compare(calendarLastDate, theDate) === 1) {
        lastMasterTaskID += 1;
        let endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        let frozenDate = theDate.clone();
        while (!workingDatesStr.includes(endDateStr)) {
          theDate = Date.parse(theDate).addDays(-1);
          endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        }
        if (skipSundays === "Yes" && theDate.is().sunday()) {
          theDate = Date.parse(theDate).addDays(-1);
        }
        masterArray.push([theDoer, lastMasterTaskID, theFreq, theTask, new Date(theDate)]);
        theDate = Date.parse(frozenDate).addMonths(1);
      }

    // ---- HALF-YEARLY (H) ----
    } else if (theFreq === "H") {
      while (Date.compare(calendarLastDate, theDate) === 1) {
        lastMasterTaskID += 1;
        let endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        let frozenDate = theDate.clone();

        while (!workingDatesStr.includes(endDateStr)) {
          theDate = Date.parse(theDate).addDays(-1);
          endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        }

        if (skipSundays === "Yes" && theDate.is().sunday()) {
          theDate = Date.parse(theDate).addDays(-1);
          endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1);
            endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
          }
        }

        masterArray.push([theDoer, lastMasterTaskID, theFreq, theTask, new Date(theDate)]);
        theDate = Date.parse(frozenDate).addMonths(6);
      }

    // ---- YEARLY (Y) ----
    } else if (theFreq === "Y") {
      lastMasterTaskID += 1;
      theDate = Date.parse(theDate).addYears(1);
      masterArray.push([theDoer, lastMasterTaskID, theFreq, theTask, new Date(theDate)]);

    // ---- DAILY (D) ----
    } else if (theFreq === "D") {
      while (Date.compare(calendarLastDate, theDate) === 1) {
        lastMasterTaskID += 1;
        let endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        let frozenDate = theDate.clone();

        if (skipSundays === "Yes") {
          if (workingDatesStr.includes(endDateStr) && !theDate.is().sunday()) {
            masterArray.push([theDoer, lastMasterTaskID, theFreq, theTask, new Date(theDate)]);
          }
        } else {
          if (workingDatesStr.includes(endDateStr)) {
            masterArray.push([theDoer, lastMasterTaskID, theFreq, theTask, new Date(theDate)]);
          }
        }

        theDate = Date.parse(frozenDate).addDays(1);
      }

    // ---- QUARTERLY (Q) ----
    } else if (theFreq === "Q") {
      while (Date.compare(calendarLastDate, theDate) === 1) {
        lastMasterTaskID += 1;
        let endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        let frozenDate = theDate.clone();

        while (!workingDatesStr.includes(endDateStr)) {
          theDate = Date.parse(theDate).addDays(-1);
          endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        }

        if (skipSundays === "Yes" && theDate.is().sunday()) {
          theDate = Date.parse(theDate).addDays(-1);
        }

        masterArray.push([theDoer, lastMasterTaskID, theFreq, theTask, new Date(theDate)]);
        theDate = Date.parse(frozenDate).addMonths(3);
      }

    // ---- FORTNIGHTLY (F) ----
    } else if (theFreq === "F") {
      while (Date.compare(calendarLastDate, theDate) === 1) {
        lastMasterTaskID += 1;
        let endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        let frozenDate = theDate.clone();

        while (!workingDatesStr.includes(endDateStr)) {
          theDate = Date.parse(theDate).addDays(-1);
          endDateStr = Utilities.formatDate(theDate, tz, "yyyy-MM-dd");
        }

        masterArray.push([theDoer, lastMasterTaskID, theFreq, theTask, new Date(theDate)]);
        theDate = Date.parse(frozenDate).addWeeks(2);
      }
    }

    // Mark row as sent
    taskListSheet.getRange(index + 2, 5).setValue("Sent");
  });

  // Bulk write to Master
  if (masterArray.length > 0) {
    masterSheet.getRange(masterLastRow + 1, 1, masterArray.length, 5).setValues(masterArray);
  }
}

// Timestamp "Done" in Master (col just left to the cell edited to "Done")
function onEdit(event) {
  if (!event) return;
  if (event.value === "Done") {
    const sheet = event.range.getSheet();
    if (sheet.getName() === "Master") {
      const row = event.range.getRow();
      const col = event.range.getColumn();
      const tz = Session.getScriptTimeZone();
      const dateStr = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");

      const ts = sheet.getRange(row, col - 1).getValue(); // Actual Time column
      if (!ts) {
        sheet.getRange(row, col - 1).setValue(dateStr);
      }
    }
  }
}
