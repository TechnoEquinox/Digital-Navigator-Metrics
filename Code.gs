// Created by Connor Bailey
// Wallingford Public Library: Tech Connect Program
// Version 1.1
//
// About:
// This script analyzes the Appointment Summary (Responses) sheet to provide metrics for
// Tech Connect employees. This script currently calculates the total appointments seen
// by the entire department for the month, quarter, and year. Additionally, this script
// calculates an individual employees total appointments for the month, quarter, and year.
// All of these statistics are then feed into a horizontal bar graph that visualizes the
// appointments over the year.

// Counts the total appointments for the current month
function getMonthlyAppointments() {
  return countAppointments(new Date().getFullYear(), new Date().getMonth());
}

// Counts the total appointments for the current quarter
function getQuarterlyAppointments() {
  return countAppointments(new Date().getFullYear(), getQuarterMonths());
}

// Counts the total appointments for the current year
function getYearlyAppointments() {
  return countAppointments(new Date().getFullYear());
}

// Counts appointments for a specific month, quarter, or year
function countAppointments(year, monthOrMonths) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var values = sheet.getDataRange().getValues();
  var count = 0;

  values.slice(1).forEach(row => { // Skip header row
    var appointmentDate = new Date(row[0]);
    if (isValidDate(appointmentDate) && isWithinPeriod(appointmentDate, year, monthOrMonths)) {
      count++;
    }
  });

  return count;
}

// Checks if a date falls within the specified year and month(s)
function isWithinPeriod(date, year, monthOrMonths) {
  return date.getFullYear() === year && 
         (Array.isArray(monthOrMonths) 
           ? monthOrMonths.includes(date.getMonth()) 
           : date.getMonth() === monthOrMonths);
}

// Returns an array of counts of appointments by month for the current year
function getAppointmentsByMonth() {
  var appointmentsByMonth = Array(12).fill(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1").getDataRange().getValues().slice(1).forEach(row => {
    var date = new Date(row[0]);
    if (isValidDate(date) && date.getFullYear() === new Date().getFullYear()) {
      appointmentsByMonth[date.getMonth()]++;
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
  return date.toString() !== "Invalid Date";
}

// Gets the current month name
function getCurrentMonthName() {
  return ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"][new Date().getMonth()];
}

function doGet() {
  // Start measuring execution time
  var startTime = new Date().getTime();

  // General stats
  var currentYear = new Date().getFullYear();
  var currentQuarter = getQuarterMonths();
  var monthlyCount = getMonthlyAppointments();
  var quarterlyCount = getQuarterlyAppointments();
  var yearlyCount = getYearlyAppointments();
  var appointmentsByMonth = getAppointmentsByMonth();

  // Employee stats
  // NOTE: Input Employee names here
  var employees = ["Employee 1", "Employee 2", "Employee 3"];
  
  var employeeStats = employees.map(employeeName => ({
    name: employeeName,
    monthlyCount: getEmployeeAppointments(employeeName, { year: currentYear, monthOrMonths: new Date().getMonth() }),
    quarterlyCount: getEmployeeAppointments(employeeName, { year: currentYear, monthOrMonths: currentQuarter }),
    yearlyCount: getEmployeeAppointments(employeeName, { year: currentYear }),
    appointmentsByMonth: getAppointmentsByMonthForEmployee(employeeName)
  }));

  // End measuring execution time
  var endTime = new Date().getTime();
  var executionTime = (endTime - startTime) / 1000; // in seconds

  // Setup HTML template and pass data
  var template = HtmlService.createTemplateFromFile('index');
  template.monthlyCount = monthlyCount;
  template.quarterlyCount = quarterlyCount;
  template.yearlyCount = yearlyCount;
  template.currentQuarter = "Q" + (Math.floor(new Date().getMonth() / 3) + 1);
  template.currentMonthName = getCurrentMonthName();
  template.currentYear = currentYear;
  template.appointmentsByMonth = appointmentsByMonth;
  template.employeeStats = JSON.stringify(employeeStats);
  template.executionTime = executionTime.toFixed(3);

  return template.evaluate().setTitle("Tech Connect Metrics (WIP)");
}
