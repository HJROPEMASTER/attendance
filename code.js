// Enhanced Attendance System with Improved Formula Synchronization

// Global cache for frequently accessed data
var CACHE = {
  employeeIds: null,
  lastCacheUpdate: null,
  CACHE_DURATION: 5 * 60 * 1000 // 5 minutes
};

function doGet() {
  return HtmlService.createTemplateFromFile("index").evaluate()
      .setTitle("HJ Rope Master Inc. - Time Tracker")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Function to get the status of clock-in/out from the \'Settings\' sheet
function getClockStatus() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName("Settings");
    if (!settingsSheet) {
      // Create and initialize the Settings sheet if it doesn\'t exist
      settingsSheet = ss.insertSheet("Settings");
      settingsSheet.getRange("A1:B1").setValues([["Setting", "Status"]]).setFontWeight("bold");
      settingsSheet.getRange("A2:B2").setValues([["ClockIn", "OFF"]]);
      settingsSheet.getRange("A3:B3").setValues([["ClockOut", "ON"]]); // Changed to ON by default
    }
    var settings = settingsSheet.getRange("A2:B3").getValues();
    var status = {};
    for (var i = 0; i < settings.length; i++) {
      status[settings[i][0]] = settings[i][1];
    }
    return status;
  } catch (e) {
    console.error("Error in getClockStatus:", e);
    // Return a default status in case of an error
    return { ClockIn: 'OFF', ClockOut: 'ON' };
  }
}

// Enhanced clockIn with formula synchronization
function clockIn(employee, gps) {
  try {
    var location = "Unknown Location";
    var newDate = new Date();

    // Get location if GPS coordinates are provided
    if (gps && gps[0] !== 0 && gps[1] !== 0) {
      try {
        var response = Maps.newGeocoder().setRegion("PHIL").setLanguage("en-PH").reverseGeocode(gps[0], gps[1]);
        location = response.results[0]?.formatted_address || "Unknown Location";
      } catch (e) {
        console.log("Location service unavailable, using default location");
        location = "Location unavailable";
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("MAIN");

    // Ensure the sheet exists and has proper headers
    if (!mainSheet) {
      mainSheet = ss.insertSheet("MAIN");
      var headers = ["Employee ID", "Clock In", "Clock Out", "Clock In Location", "Clock Out Location", "Total Hours"];
      mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      mainSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    // Batch read all data at once for better performance
    var dataRange = mainSheet.getDataRange();
    var allData = dataRange.getValues();
    var lastRow = allData.length;

    var returnDate = formatDate(newDate);
    var msg = "SUCCESS";
    var returnArray = [];

    // Check for existing clock-in (more efficient loop)
    // Iterate backwards to find the most recent un-clocked-out entry quickly
    for (var j = allData.length - 1; j >= 1; j--) {
      if (employee === allData[j][0] && !allData[j][2]) { // Column A (employee) and Column C (clock out)
        msg = "Pasensya na, kailangan mo munang mag-ClockOut!" + " " + "Na-ClockIn ka na kanina sa oras na: " + formatDate(allData[j][1]);
        returnArray.push([msg, returnDate, employee]);
        return returnArray;
      }
    }

    // Add new clock-in record
    var newRowData = [employee, newDate, "", location, "", ""];
    mainSheet.appendRow(newRowData); // Use appendRow for single row addition, then format

    // Apply formatting in batch for the newly added row
    var newRowRange = mainSheet.getRange(lastRow + 1, 1, 1, 6);
    newRowRange.setFontSize(10);
    mainSheet.getRange(lastRow + 1, 2).setNumberFormat("dd/MM/yyyy - HH:mm:ss").setHorizontalAlignment("left");

    // Trigger formula synchronization
    synchronizeFormulas();

    returnArray.push([msg, returnDate, employee]);
    return returnArray;

  } catch (error) {
    console.error("Error in clockIn:", error);
    return [["Error occurred during clock in. Please try again.", formatDate(new Date()), employee]];
  }
}

// Enhanced clockOut with formula synchronization
function clockOut(employee, gps) {
  try {
    var location = "Unknown Location";
    var newDate = new Date();

    // Get location if GPS coordinates are provided
    if (gps && gps[0] !== 0 && gps[1] !== 0) {
      try {
        var response = Maps.newGeocoder().setRegion("PHIL").setLanguage("en-PH").reverseGeocode(gps[0], gps[1]);
        location = response.results[0]?.formatted_address || "Unknown Location";
      } catch (e) {
        console.log("Location service unavailable, using default location");
        location = "Location unavailable";
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("MAIN");

    if (!mainSheet) {
      return [["Error: MAIN sheet not found. Please clock in first.", formatDate(new Date()), employee]];
    }

    // Batch read all data at once for better performance
    var dataRange = mainSheet.getDataRange();
    var allData = dataRange.getValues();

    var returnDate = formatDate(newDate);
    var msg = "SUCCESS";
    var foundRecord = false;
    var returnArray = [];

    // Find and update the clock-out record (more efficient - iterate backwards)
    for (var j = allData.length - 1; j >= 1; j--) {
      if (employee === allData[j][0] && !allData[j][2]) { // Column A (employee) and Column C (clock out)
        // Preserve the Clock In Location (Column D) before updating the row
        var clockInLocation = allData[j][3];

        // Calculate total time
        var totalTime = (newDate - allData[j][1]) / (60 * 60 * 1000);

        // Update the row with clock-out data, including the preserved Clock In Location
        var updateRange = mainSheet.getRange(j + 1, 3, 1, 4); // Columns C to F
        updateRange.setValues([[newDate, clockInLocation, location, totalTime.toFixed(2)]]);

        // Apply formatting
        mainSheet.getRange(j + 1, 3).setNumberFormat("MM/dd/yyyy - HH:mm:ss").setHorizontalAlignment("left").setFontSize(10);
        mainSheet.getRange(j + 1, 5).setFontSize(10);
        mainSheet.getRange(j + 1, 6).setNumberFormat("#0.00").setHorizontalAlignment("left").setFontSize(12);

        foundRecord = true;
        break;
      }
    }

    if (!foundRecord) {
      returnArray.push(["Pasensya na, kailangan mo munong mag-ClockIn!, Naka-ClockOut ka na kanina:", "", employee]);
      return returnArray;
    }

    // Trigger comprehensive formula synchronization after successful clock-out
    synchronizeFormulas();

    returnArray.push([msg, returnDate, employee]);
    return returnArray;

  } catch (error) {
    console.error("Error in clockOut:", error);
    return [["Error occurred during clock out. Please try again.", formatDate(new Date()), employee]];
  }
}

// Enhanced formula synchronization function
function synchronizeFormulas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Synchronize main calculations
    calculateTotalHours();
    updateDailySummary();
    updateWeeklySummary();
    updateMonthlySummary();

    // Force recalculation of all formulas in the spreadsheet
    SpreadsheetApp.flush();

    console.log("Formula synchronization completed successfully");

  } catch (error) {
    console.error("Error in synchronizeFormulas:", error);
  }
}

// Enhanced calculateTotalHours function with better formula management
function calculateTotalHours() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("MAIN");

    if (!mainSheet) {
      console.log("MAIN sheet not found, skipping total hours calculation");
      return;
    }

    // Read all data at once for better performance
    var dataRange = mainSheet.getDataRange();
    var allData = dataRange.getValues();
    var totals = {};

    // Process data more efficiently using object for totals
    for (var j = 1; j < allData.length; j++) { // Start from 1 to skip header
      var rate = allData[j][5]; // Column F (total hours)
      var name = allData[j][0]; // Column A (employee name)

      if (rate && name) {
        if (totals[name]) {
          totals[name] += parseFloat(rate);
        } else {
          totals[name] = parseFloat(rate);
        }
      }
    }

    // Clear existing totals and write new ones
    var summaryRange = mainSheet.getRange("H2:I" + mainSheet.getLastRow()); // Clear only existing data
    summaryRange.clearContent();

    // Add headers for summary section
    mainSheet.getRange("H1").setValue("Employee Summary");
    mainSheet.getRange("I1").setValue("Total Hours");
    mainSheet.getRange("H1:I1").setFontWeight("bold");

    // Convert object to array and write in batch
    var totalArray = [];
    for (var name in totals) {
      totalArray.push([name, totals[name]]);
    }

    if (totalArray.length > 0) {
      var outputRange = mainSheet.getRange(2, 8, totalArray.length, 2); // Columns H and I
      outputRange.setValues(totalArray);
      outputRange.setFontSize(12);

      // Format the total hours column
      mainSheet.getRange(2, 9, totalArray.length, 1).setNumberFormat("#0.00");
    }

  } catch (error) {
    console.error("Error in calculateTotalHours:", error);
  }
}

// New function to update daily summary
function updateDailySummary() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("MAIN");
    var summarySheet = ss.getSheetByName("DAILY SUMMARY");

    if (!summarySheet) {
      summarySheet = ss.insertSheet("DAILY SUMMARY");
      var headers = ["Date", "Employee ID", "Clock In", "Clock Out", "Total Hours", "Status"];
      summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      summarySheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    if (!mainSheet) return;

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var dataRange = mainSheet.getDataRange();
    var allData = dataRange.getValues();
    var todayRecords = [];

    // Filter today\'s records
    for (var i = 1; i < allData.length; i++) {
      var recordDate = new Date(allData[i][1]);
      recordDate.setHours(0, 0, 0, 0);

      if (recordDate.getTime() === today.getTime()) {
        var status = allData[i][2] ? "Complete" : "Clock In Only";
        todayRecords.push([
          formatDate(allData[i][1]).split(" - ")[0], // Date only
          allData[i][0], // Employee ID
          formatDate(allData[i][1]), // Clock In
          allData[i][2] ? formatDate(allData[i][2]) : "Not clocked out", // Clock Out
          allData[i][5] || 0, // Total Hours
          status
        ]);
      }
    }

    // Clear existing data and write new summary
    if (summarySheet.getLastRow() > 1) {
      summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 6).clearContent(); // Clear only content, not formatting
    }
    if (todayRecords.length > 0) {
      summarySheet.getRange(2, 1, todayRecords.length, 6).setValues(todayRecords);
    }

  } catch (error) {
    console.error("Error in updateDailySummary:", error);
  }
}

// New function to update weekly summary
function updateWeeklySummary() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("MAIN");
    var weeklySheet = ss.getSheetByName("WEEKLY SUMMARY");

    if (!weeklySheet) {
      weeklySheet = ss.insertSheet("WEEKLY SUMMARY");
      var headers = ["Week Starting", "Employee ID", "Total Hours", "Days Worked", "Average Hours/Day"];
      weeklySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      weeklySheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    if (!mainSheet) return;

    var dataRange = mainSheet.getDataRange();
    var allData = dataRange.getValues();
    var weeklyTotals = {};

    // Get start of current week (Monday)
    var today = new Date();
    var dayOfWeek = today.getDay();
    var mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // Handle Sunday as 0
    var monday = new Date(today);
    monday.setDate(today.getDate() + mondayOffset);
    monday.setHours(0, 0, 0, 0);

    // Process this week\'s data
    for (var i = 1; i < allData.length; i++) {
      var recordDate = new Date(allData[i][1]);

      if (recordDate >= monday && allData[i][5]) {
        var employee = allData[i][0];
        var hours = parseFloat(allData[i][5]);

        if (!weeklyTotals[employee]) {
          weeklyTotals[employee] = { totalHours: 0, daysWorked: 0, dates: new Set() };
        }

        weeklyTotals[employee].totalHours += hours;

        // Count unique days
        var dateKey = recordDate.toDateString();
        if (!weeklyTotals[employee].dates.has(dateKey)) {
          weeklyTotals[employee].dates.add(dateKey);
          weeklyTotals[employee].daysWorked++;
        }
      }
    }

    // Clear existing data and write new summary
    if (weeklySheet.getLastRow() > 1) {
      weeklySheet.getRange(2, 1, weeklySheet.getLastRow() - 1, 5).clearContent();
    }

    var weeklyArray = [];
    var weekStartStr = formatDate(monday).split(" - ")[0];

    for (var employee in weeklyTotals) {
      var data = weeklyTotals[employee];
      var avgHours = data.daysWorked > 0 ? (data.totalHours / data.daysWorked).toFixed(2) : 0;

      weeklyArray.push([
        weekStartStr,
        employee,
        data.totalHours.toFixed(2),
        data.daysWorked,
        avgHours
      ]);
    }

    if (weeklyArray.length > 0) {
      weeklySheet.getRange(2, 1, weeklyArray.length, 5).setValues(weeklyArray);
    }

  } catch (error) {
    console.error("Error in updateWeeklySummary:", error);
  }
}

// New function to update monthly summary
function updateMonthlySummary() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("MAIN");
    var monthlySheet = ss.getSheetByName("MONTHLY SUMMARY");

    if (!monthlySheet) {
      monthlySheet = ss.insertSheet("MONTHLY SUMMARY");
      var headers = ["Month/Year", "Employee ID", "Total Hours", "Days Worked", "Average Hours/Day"];
      monthlySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      monthlySheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    if (!mainSheet) return;

    var dataRange = mainSheet.getDataRange();
    var allData = dataRange.getValues();
    var monthlyTotals = {};

    var today = new Date();
    var currentMonth = today.getMonth();
    var currentYear = today.getFullYear();

    // Process this month\'s data
    for (var i = 1; i < allData.length; i++) {
      var recordDate = new Date(allData[i][1]);

      if (recordDate.getMonth() === currentMonth &&
          recordDate.getFullYear() === currentYear &&
          allData[i][5]) {

        var employee = allData[i][0];
        var hours = parseFloat(allData[i][5]);

        if (!monthlyTotals[employee]) {
          monthlyTotals[employee] = { totalHours: 0, daysWorked: 0, dates: new Set() };
        }

        monthlyTotals[employee].totalHours += hours;

        // Count unique days
        var dateKey = recordDate.toDateString();
        if (!monthlyTotals[employee].dates.has(dateKey)) {
          monthlyTotals[employee].dates.add(dateKey);
          monthlyTotals[employee].daysWorked++;
        }
      }
    }

    // Clear existing data and write new summary
    if (monthlySheet.getLastRow() > 1) {
      monthlySheet.getRange(2, 1, monthlySheet.getLastRow() - 1, 5).clearContent();
    }

    var monthlyArray = [];
    var monthNames = ["January", "February", "March", "April", "May", "June",
                     "July", "August", "September", "October", "November", "December"];
    var monthYearStr = monthNames[currentMonth] + " " + currentYear;

    for (var employee in monthlyTotals) {
      var data = monthlyTotals[employee];
      var avgHours = data.daysWorked > 0 ? (data.totalHours / data.daysWorked).toFixed(2) : 0;

      monthlyArray.push([
        monthYearStr,
        employee,
        data.totalHours.toFixed(2),
        data.daysWorked,
        avgHours
      ]);
    }

    if (monthlyArray.length > 0) {
      monthlySheet.getRange(2, 1, monthlyArray.length, 5).setValues(monthlyArray);
    }

  } catch (error) {
    console.error("Error in updateMonthlySummary:", error);
  }
}

function formatDate(date) {
  // Ensure the input is a valid Date object before formatting
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    console.warn("Invalid date provided to formatDate:", date);
    return "Invalid Date"; // Or handle as appropriate for your application
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy - HH:mm:ss");
}

// Enhanced employee ID retrieval with better caching
function getEmployeeIDs() {
  try {
    var now = new Date().getTime();

    // Return cached data if still valid
    if (CACHE.employeeIds && CACHE.lastCacheUpdate &&
        (now - CACHE.lastCacheUpdate) < CACHE.CACHE_DURATION) {
      return CACHE.employeeIds;
    }

    // Fetch fresh data
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Employee Name");

    if (!sheet) {
      // Create the sheet if it doesn\'t exist
      sheet = ss.insertSheet("Employee Name");
      var headers = ["Employee Name", "Employee ID"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

      // Add some sample data
      var sampleData = [
        ["John Doe", "EMP001"],
        ["Jane Smith", "EMP002"],
        ["Mike Johnson", "EMP003"]
      ];
      sheet.getRange(2, 1, sampleData.length, 2).setValues(sampleData);

      console.log("Created Employee Name sheet with sample data");
    }

    var range = sheet.getRange("B2:B" + sheet.getLastRow()); // Read only existing data
    var values = range.getValues().flat().filter(id => id && id.toString().trim() !== "");

    // Update cache
    CACHE.employeeIds = values;
    CACHE.lastCacheUpdate = now;

    return values;

  } catch (error) {
    console.error("Error in getEmployeeIDs:", error);
    return [];
  }
}

// New function to get attendance statistics for HTML display
function getAttendanceStats() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("MAIN");

    if (!mainSheet) {
      return {
        totalRecords: 0,
        todayRecords: 0,
        activeEmployees: 0,
        lastUpdate: formatDate(new Date())
      };
    }

    var dataRange = mainSheet.getDataRange();
    var allData = dataRange.getValues();

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var stats = {
      totalRecords: allData.length > 1 ? allData.length - 1 : 0, // Exclude header
      todayRecords: 0,
      activeEmployees: new Set(),
      lastUpdate: formatDate(new Date())
    };

    for (var i = 1; i < allData.length; i++) {
      var recordDate = new Date(allData[i][1]);
      recordDate.setHours(0, 0, 0, 0);

      if (recordDate.getTime() === today.getTime()) {
        stats.todayRecords++;
      }

      if (allData[i][0]) {
        stats.activeEmployees.add(allData[i][0]);
      }
    }

    stats.activeEmployees = stats.activeEmployees.size;

    return stats;

  } catch (error) {
    console.error("Error in getAttendanceStats:", error);
    return {
      totalRecords: 0,
      todayRecords: 0,
      activeEmployees: 0,
      lastUpdate: formatDate(new Date()),
      error: error.toString()
    };
  }
}

// New function to get employee\'s current status
function getEmployeeStatus(employeeId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("MAIN");

    if (!mainSheet) {
      return { status: "no_records", message: "No attendance data found" };
    }

    var dataRange = mainSheet.getDataRange();
    var allData = dataRange.getValues();

    // Find the most recent record for this employee
    for (var i = allData.length - 1; i >= 1; i--) {
      if (allData[i][0] === employeeId) {
        if (allData[i][2]) {
          // Has clock out time
          return {
            status: "clocked_out",
            message: "Last clocked out at " + formatDate(allData[i][2]),
            lastClockIn: formatDate(allData[i][1]),
            lastClockOut: formatDate(allData[i][2]),
            totalHours: allData[i][5] || 0
          };
        } else {
          // Only has clock in time
          return {
            status: "clocked_in",
            message: "Currently clocked in since " + formatDate(allData[i][1]),
            lastClockIn: formatDate(allData[i][1]),
            lastClockOut: null,
            totalHours: 0
          };
        }
      }
    }

    return { 
      status: "no_records",
      message: "No attendance records found for this employee"
    };

  } catch (error) {
    console.error("Error in getEmployeeStatus:", error);
    return {
      status: "error",
      message: "Error retrieving employee status: " + error.toString()
    };
  }
}

// Enhanced email function with better error handling
function sendAttendanceReport() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DAILY SUMMARY");

    if (!sheet) {
      updateDailySummary(); // Create the summary if it doesn\'t exist
      sheet = ss.getSheetByName("DAILY SUMMARY");
    }

    var data = sheet.getDataRange().getValues();
    var reportContent = "Daily Attendance Report - " + formatDate(new Date()).split(" - ")[0] + "\n\n";

    for (var i = 1; i < data.length; i++) {
      var date = data[i][0];
      var employee = data[i][1];
      var clockIn = data[i][2];
      var clockOut = data[i][3];
      var totalHours = data[i][4];
      var status = data[i][5];

      reportContent += `${employee}: ${clockIn} - ${clockOut} (${totalHours} hours) [${status}]\n`;
    }

    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter your email address", "Please enter your email:", ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() === ui.Button.OK) {
      var email = response.getResponseText().trim();
      if (email && /^[\S+@\S+\.\S+]$/i.test(email)) { // Corrected regex
        MailApp.sendEmail({
          to: email,
          subject: "Daily Attendance Report - " + formatDate(new Date()).split(" - ")[0],
          body: reportContent,
        });
        ui.alert(`Report sent successfully to ${email}`);
      } else {
        ui.alert("Invalid email address. Please try again.");
      }
    } else {
      ui.alert("Operation canceled.");
    }

  } catch (error) {
    console.error("Error in sendAttendanceReport:", error);
    SpreadsheetApp.getUi().alert("Error sending report. Please try again.");
  }
}
