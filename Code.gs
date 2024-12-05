// Counts the total appointments for the current month
function getMonthlyAppointments() {
  var result = countAppointments(new Date().getFullYear(), new Date().getMonth());
  return { total: result.count, locationCounts: result.locationCounts };
}

// Counts the total appointments for the current quarter
function getQuarterlyAppointments() {
  var result = countAppointments(new Date().getFullYear(), getQuarterMonths());
  return { total: result.count, locationCounts: result.locationCounts };
}

// Counts the total appointments for the current year
function getYearlyAppointments() {
  var result = countAppointments(new Date().getFullYear(), undefined);
  return { total: result.count, locationCounts: result.locationCounts };
}

// Counts appointments for a specific month, quarter, or year
function countAppointments(year, monthOrMonths) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var values = sheet.getDataRange().getValues();
  var count = 0;

  // Initialize location counts
  locationCounts = {
    Library: 0,
    SCOW: 0,
    SeniorCenter: 0,
    MastersManna: 0,
    Other: 0
  };
  
  Logger.log("Initialized Location Counts: " + JSON.stringify(locationCounts));

  values.slice(1).forEach(row => { // Skip header row
    var appointmentDate = new Date(row[0]);
    var location = row[4] ? row[4].trim() : "Other"; // Normalize location and default to "Other"

    if (isValidDate(appointmentDate) && isWithinPeriod(appointmentDate, year, monthOrMonths)) {
      count++;

      switch (location) {
        case "Library":
          locationCounts.Library++;
          break;
        case "SCOW":
          locationCounts.SCOW++;
          break;
        case "Senior Center":
          locationCounts.SeniorCenter++;
          break;
        case "Master's Manna":
          locationCounts.MastersManna++;
          break;
        default:
          locationCounts.Other++;
          break;
      }
    }
  });

  Logger.log("Final Location Counts: " + JSON.stringify(locationCounts));
  return { count, locationCounts };
}

// Checks if a date falls within the specified year and month(s)
// Adjusted to handle yearly counting when monthOrMonths is not provided
function isWithinPeriod(date, year, monthOrMonths) {
  if (date.getFullYear() !== year) return false;
  
  // If monthOrMonths is undefined, count the entire year
  if (monthOrMonths === undefined) return true;
  
  // Handle both single month and array of months (e.g., quarter)
  return Array.isArray(monthOrMonths)
    ? monthOrMonths.includes(date.getMonth())
    : date.getMonth() === monthOrMonths;
}

// Returns an array of counts of appointments by month for the current year up to the current date
function getAppointmentsByMonth() {
  var appointmentsByMonth = Array(12).fill(0);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var values = sheet.getDataRange().getValues().slice(1); // Skip header row
  var currentDate = new Date();
  var currentYear = currentDate.getFullYear();
  var currentMonth = currentDate.getMonth();
  var currentDay = currentDate.getDate();

  values.forEach(row => {
    var date = new Date(row[0]);
    if (isValidDate(date) && date.getFullYear() === currentYear) {
      // Only count dates up to the current month and current day
      if (
        date.getMonth() < currentMonth || 
        (date.getMonth() === currentMonth && date.getDate() <= currentDay)
      ) {
        appointmentsByMonth[date.getMonth()]++;
      }
    }
  });

  return appointmentsByMonth;
}


// Counts monthly, quarterly, or yearly appointments for a specific employee
function getEmployeeAppointments(employeeName, period) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var values = sheet.getDataRange().getValues();
  var count = 0;

  values.slice(1).forEach(row => {
    var appointmentDate = new Date(row[0]);
    var navigatorName = row[1];
    if (isValidDate(appointmentDate) && navigatorName.includes(employeeName) && isWithinPeriod(appointmentDate, period.year, period.monthOrMonths)) {
      count++;
    }
  });

  return count;
}

// Returns an array of monthly appointment counts for a specific employee for the current year
function getAppointmentsByMonthForEmployee(employeeName) {
  var appointmentsByMonth = Array(12).fill(0);
  var currentYear = new Date().getFullYear();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1").getDataRange().getValues().slice(1).forEach(row => {
    var date = new Date(row[0]);
    if (isValidDate(date) && date.getFullYear() === currentYear && row[1] === employeeName) {
      appointmentsByMonth[date.getMonth()]++;
    }
  });
  return appointmentsByMonth;
}

// Returns an array representing the months in the current quarter
function getQuarterMonths() {
  var currentMonth = new Date().getMonth();
  if (currentMonth <= 2) return [0, 1, 2];      // Q1
  if (currentMonth <= 5) return [3, 4, 5];      // Q2
  if (currentMonth <= 8) return [6, 7, 8];      // Q3
  return [9, 10, 11];                           // Q4
}

// Validates if a given date object is valid
function isValidDate(date) {
  return date.toString() !== "Invalid Date" && !isNaN(date.getTime());
}

// Gets the current month name
function getCurrentMonthName() {
  return ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"][new Date().getMonth()];
}

// Helper function to calculate average appointment times for a given period
function getAverageAppointmentTime(year, monthOrMonths) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var values = sheet.getDataRange().getValues();
  var totalMinutes = 0;
  var count = 0;

  values.slice(1).forEach(row => { // Skip header row
    var appointmentDate = new Date(row[0]);
    var appointmentTime = row[5];
    if (isValidDate(appointmentDate) && isWithinPeriod(appointmentDate, year, monthOrMonths) && appointmentTime > 0) {
      totalMinutes += parseInt(appointmentTime);
      count++;
    }
  });

  return count > 0 ? totalMinutes / count : 0;
}

// Helper function to calculate average appointment times for an employee for a given period
function getEmployeeAverageAppointmentTime(employeeName, period) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var values = sheet.getDataRange().getValues();
  var totalMinutes = 0;
  var count = 0;

  values.slice(1).forEach(row => {
    var appointmentDate = new Date(row[0]);
    var navigatorName = row[1];
    var appointmentTime = row[5]; 
    if (isValidDate(appointmentDate) && navigatorName.includes(employeeName) && isWithinPeriod(appointmentDate, period.year, period.monthOrMonths) && appointmentTime > 0) {
      totalMinutes += parseInt(appointmentTime);
      count++;
    }
  });

  return count > 0 ? totalMinutes / count : 0;
}

function doGet() {
  // Start measuring execution time
  var startTime = new Date().getTime();

  // General stats
  var currentYear = new Date().getFullYear();
  var currentQuarter = getQuarterMonths();

  // Calculate totals and location counts
  var monthlyData = getMonthlyAppointments();
  var quarterlyData = getQuarterlyAppointments();
  var yearlyData = getYearlyAppointments();

  var locationCounts = yearlyData.locationCounts; // Using yearly data for total location count

  // DEBUG
  Logger.log("Library Appointments: " + locationCounts)

  var appointmentsByMonth = getAppointmentsByMonth();
  var averageMonthlyTime = getAverageAppointmentTime(currentYear, new Date().getMonth());
  var averageQuarterlyTime = getAverageAppointmentTime(currentYear, currentQuarter);
  var averageYearlyTime = getAverageAppointmentTime(currentYear);

  // Employee stats
  var employees = ["Connor Bailey", "Elijah Mitchell"];
  
  var employeeStats = employees.map(employeeName => ({
    name: employeeName,
    monthlyCount: getEmployeeAppointments(employeeName, { year: currentYear, monthOrMonths: new Date().getMonth() }),
    quarterlyCount: getEmployeeAppointments(employeeName, { year: currentYear, monthOrMonths: currentQuarter }),
    yearlyCount: getEmployeeAppointments(employeeName, { year: currentYear }),
    appointmentsByMonth: getAppointmentsByMonthForEmployee(employeeName),
    averageMonthlyTime: getEmployeeAverageAppointmentTime(employeeName, { year: currentYear, monthOrMonths: new Date().getMonth() }),
    averageQuarterlyTime: getEmployeeAverageAppointmentTime(employeeName, { year: currentYear, monthOrMonths: currentQuarter }),
    averageYearlyTime: getEmployeeAverageAppointmentTime(employeeName, { year: currentYear })
  }));

  // End measuring execution time
  var endTime = new Date().getTime();
  var executionTime = (endTime - startTime) / 1000; // in seconds

  // Setup HTML template and pass data
  var template = HtmlService.createTemplateFromFile('index');
  template.monthlyCount = monthlyData.total;
  template.quarterlyCount = quarterlyData.total;
  template.yearlyCount = yearlyData.total;
  template.locationCounts = JSON.stringify(locationCounts);

  template.currentQuarter = "Q" + (Math.floor(new Date().getMonth() / 3) + 1);
  template.currentMonthName = getCurrentMonthName();
  template.currentYear = currentYear;
  template.appointmentsByMonth = appointmentsByMonth;
  template.employeeStats = JSON.stringify(employeeStats);
  template.executionTime = executionTime.toFixed(3);
  template.averageMonthlyTime = averageMonthlyTime;
  template.averageQuarterlyTime = averageQuarterlyTime;
  template.averageYearlyTime = averageYearlyTime;

  return template.evaluate().setTitle("Tech Connect Metrics");
}
