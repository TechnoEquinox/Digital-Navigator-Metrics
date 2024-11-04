// Created by Connor Bailey
// Wallingford Public Library: Tech Connect Program
// Version 1.0
//
// About:
// This script analyzes the Appointment Summary (Responses) sheet to provide metrics for
// Tech Connect employees. This script currently calculates the total appointments seen
// by the entire department for the month, quarter, and year. Additionally, this script
// calculates an individual employees total appointments for the month, quarter, and year.
// All of these statistics are then feed into a horizontal bar graph that visualizes the
// appointments over the year.
//
//
// Feature Request:
// - Migrate this code to GitHub
// - Calculate the average appointment time for each employee
// - Calculate the most frequently seen clients for each employee
// - Calculate the top three most popular "main goals" that clients bring for all employees
// - Calculate the number of appointments seen at each location 
//
// - Provide some statistics from the Impromptu Help Sheet
//

// Function to count the number of appointments this month
function countAppointmentsThisMonth() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var currentDate = new Date();
  var currentMonth = currentDate.getMonth();  // Get current month (0-11)
  var currentYear = currentDate.getFullYear();

  var count = 0;

  // Loop through the rows to count the appointments in the current month
  for (var i = 1; i < values.length; i++) { // Assumes first row of sheet is header
    var appointmentDateStr = values[i][0];  // Get the date string from the sheet
    var appointmentDate = new Date(appointmentDateStr); // Convert to Date object

    // Check if appointmentDate is a valid date
    if (appointmentDate.toString() === "Invalid Date") {
      continue;  // Skip this iteration if the date is invalid
    }

    // Check if the appointment is in the current month and year
    if (appointmentDate.getMonth() === currentMonth && appointmentDate.getFullYear() === currentYear) {
      count++;
    }
  }

  return count;
}

// Function to count all of the appointments this quarter
function countAppointmentsThisQuarter() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var currentDate = new Date();
  var currentMonth = currentDate.getMonth();  // Get the current month (0-11)
  var currentYear = currentDate.getFullYear();

  var quarterStart, quarterEnd;

  // Determine the start and end of the current quarter
  if (currentMonth <= 2) {
    // Q1: January 1 - March 31
    quarterStart = new Date(currentYear, 0, 1);  // January 1
    quarterEnd = new Date(currentYear, 2, 31);   // March 31

  } else if (currentMonth <= 5) {
    // Q2: April 1 - June 30
    quarterStart = new Date(currentYear, 3, 1);  // April 1
    quarterEnd = new Date(currentYear, 5, 30);   // June 30

  } else if (currentMonth <= 8) {
    // Q3: July 1 - September 30
    quarterStart = new Date(currentYear, 6, 1);  // July 1
    quarterEnd = new Date(currentYear, 8, 30);   // September 30

  } else {
    // Q4: October 1 - December 31
    quarterStart = new Date(currentYear, 9, 1);  // October 1
    quarterEnd = new Date(currentYear, 11, 31);  // December 31
  }

  var count = 0;

  // Loop through the rows to count the appointments in the current quarter
  for (var i = 1; i < values.length; i++) { // Assuming first row is a header
    var appointmentDateStr = values[i][0];  // Get the date string from the sheet
    var appointmentDate = new Date(appointmentDateStr); // Convert to Date object

    // Check if appointmentDate is valid
    if (appointmentDate.toString() === "Invalid Date") {
      continue;  // Skip this iteration if the date is invalid
    }

    // Check if the appointment falls within the current quarter
    if (appointmentDate >= quarterStart && appointmentDate <= quarterEnd) {
      count++;
    }
  }

  return count;  // Return the total number of appointments for the current quarter
}

// Function to count all of the appointments this year
function countAppointmentsThisYear() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1"); // Replace with your sheet name
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var currentDate = new Date();
  var currentYear = currentDate.getFullYear();  // Get the current year

  var count = 0;

  // Loop through the rows to count the appointments in the current year
  for (var i = 1; i < values.length; i++) { // Assuming first row is a header
    var appointmentDateStr = values[i][0];  // Get the date string from the sheet
    var appointmentDate = new Date(appointmentDateStr); // Convert to Date object

    // Check if appointmentDate is a valid date
    if (appointmentDate.toString() === "Invalid Date") {
      continue;  // Skip this iteration if the date is invalid
    }

    // Check if the appointment is in the current year
    if (appointmentDate.getFullYear() === currentYear) {
      count++;
    }
  }

  return count;  // Return the total number of appointments for the current year
}

// Function that counts how many appointments have happened in each month of the year
function getAppointmentsByMonth() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var currentDate = new Date();
  var currentMonth = currentDate.getMonth();  // Get the current month (0-11 for Jan-Dec)
  var currentYear = currentDate.getFullYear();  // Get the current year

  // Initialize an array to store the count of appointments per month (0-11 for Jan-Dec)
  var appointmentsByMonth = new Array(12).fill(0); // [Jan, Feb, ..., Dec]

  // Loop through the rows to count the appointments in each month of the current year
  for (var i = 1; i < values.length; i++) { // Assuming first row is a header
    var appointmentDateStr = values[i][0];  // Get the date string from the sheet
    var appointmentDate = new Date(appointmentDateStr); // Convert to Date object

    // Check if the date is valid, is in the current year, and the month is not in the future
    if (appointmentDate.toString() !== "Invalid Date" && appointmentDate.getFullYear() === currentYear && appointmentDate.getMonth() <= currentMonth) {
      var month = appointmentDate.getMonth();  // Get the month (0 = Jan, 11 = Dec)
      appointmentsByMonth[month]++;  // Increment the count for that month
    }
  }

  return appointmentsByMonth;  // Return the array with appointment counts per month
}

function getCurrentQuarter() {
  var currentDate = new Date();
  var currentMonth = currentDate.getMonth();

  // Determine the current quarter based on the month of the year
  if (currentMonth <= 2) {
    return "Q1";
  } else if (currentMonth <= 5) {
    return "Q2";
  } else if (currentMonth <= 8) {
    return "Q3";
  } else {
    return "Q4"
  }
}

// Function to count appointments for a specific employee this month
function countAppointmentsForEmployeeThisMonth(employeeName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var currentDate = new Date();
  var currentMonth = currentDate.getMonth();  // Get current month (0-11)
  var currentYear = currentDate.getFullYear();

  var count = 0;
  
  for (var i = 1; i < values.length; i++) {
    var appointmentDateStr = values[i][0];  // Get the date string from the sheet
    var navigatorName = values[i][1];  // Assuming the "Digital Navigator" column is in the 2nd column (index 1)

    var appointmentDate = new Date(appointmentDateStr); // Convert to Date object
    if (appointmentDate.toString() === "Invalid Date") {
      continue;
    }

    // Check if the appointment is in the current month and year, and matches the employee name
    if (appointmentDate.getMonth() === currentMonth && appointmentDate.getFullYear() === currentYear && navigatorName.includes(employeeName)) {
      count++;
    }
  }

  return count;
}

// Function to count appointments for a specific employee this quarter
function countAppointmentsForEmployeeThisQuarter(employeeName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var currentDate = new Date();
  var currentMonth = currentDate.getMonth();  // Get the current month (0-11)
  var currentYear = currentDate.getFullYear();

  var quarterStart, quarterEnd;

  // Determine the start and end of the current quarter
  if (currentMonth <= 2) {
    quarterStart = new Date(currentYear, 0, 1);  // January 1
    quarterEnd = new Date(currentYear, 2, 31);   // March 31
  } else if (currentMonth <= 5) {
    quarterStart = new Date(currentYear, 3, 1);  // April 1
    quarterEnd = new Date(currentYear, 5, 30);   // June 30
  } else if (currentMonth <= 8) {
    quarterStart = new Date(currentYear, 6, 1);  // July 1
    quarterEnd = new Date(currentYear, 8, 30);   // September 30
  } else {
    quarterStart = new Date(currentYear, 9, 1);  // October 1
    quarterEnd = new Date(currentYear, 11, 31);  // December 31
  }

  var count = 0;

  for (var i = 1; i < values.length; i++) {
    var appointmentDateStr = values[i][0];
    var navigatorName = values[i][1];

    var appointmentDate = new Date(appointmentDateStr); 
    if (appointmentDate.toString() === "Invalid Date") {
      continue;
    }

    // Check if the appointment is within the current quarter and matches the employee name
    if (appointmentDate >= quarterStart && appointmentDate <= quarterEnd && navigatorName.includes(employeeName)) {
      count++;
    }
  }

  return count;
}

// Function to count appointments for a specific employee this year
function countAppointmentsForEmployeeThisYear(employeeName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var currentYear = new Date().getFullYear();  // Get the current year

  var count = 0;

  for (var i = 1; i < values.length; i++) {
    var appointmentDateStr = values[i][0];
    var navigatorName = values[i][1];

    var appointmentDate = new Date(appointmentDateStr);
    if (appointmentDate.toString() === "Invalid Date") {
      continue;
    }

    // Check if the appointment is in the current year and matches the employee name
    if (appointmentDate.getFullYear() === currentYear && navigatorName.includes(employeeName)) {
      count++;
    }
  }

  return count;
}

// Function to get the number of appointments for a specific employee in each month of the current year
function getAppointmentsByMonthForEmployee(employeeName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var currentDate = new Date();
  var currentMonth = currentDate.getMonth();  // Get the current month (0-11 for Jan-Dec)
  var currentYear = currentDate.getFullYear();  // Get the current year

  // Initialize an array to store the count of appointments per month (0-11 for Jan-Dec)
  var appointmentsByMonth = new Array(12).fill(0); // [Jan, Feb, ..., Dec]

  // Loop through the rows to count the appointments for the employee in each month of the current year
  for (var i = 1; i < values.length; i++) { // Assuming first row is a header
    var appointmentDateStr = values[i][0];  // Get the date string from the sheet
    var appointmentDate = new Date(appointmentDateStr); // Convert to Date object
    var employee = values[i][1];  // Employee name is in the 2nd column (index 1)

    // Check if the date is valid, is in the current year, and the month is not in the future
    if (appointmentDate.toString() !== "Invalid Date" && appointmentDate.getFullYear() === currentYear && appointmentDate.getMonth() <= currentMonth) {
      var month = appointmentDate.getMonth();  // Get the month (0 = Jan, 11 = Dec)

      // Check if the appointment belongs to the specified employee
      if (employee === employeeName) {
        appointmentsByMonth[month]++;  // Increment the count for that month
      }
    }
  }

  return appointmentsByMonth;  // Return the array with appointment counts per month
}

// Function to get the current month name
function getCurrentMonthName() {
  var currentDate = new Date();
  var monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return monthNames[currentDate.getMonth()]; // Return the current month name
}

function doGet() {
  var monthlyCount = countAppointmentsThisMonth(); // Total monthly count
  var quarterlyCount = countAppointmentsThisQuarter(); // Total quarterly count
  var yearlyCount = countAppointmentsThisYear(); // Total yearly count
  var currentQuarter = getCurrentQuarter(); // Current quarter
  var currentMonthName = getCurrentMonthName(); // Current month name
  var currentYear = new Date().getFullYear(); // Current year
  var appointmentsByMonth = getAppointmentsByMonth(); // Total appointments by month
  
  // Array of employee names
  // 
  // TODO: Move this to its own simple function call
  var employees = ["Connor Bailey", "Elijah Mitchell"];
  
  // Object to store the counts for each employee
  var employeeStats = [];

  // Loop through each employee and calculate their counts
  for (var i = 0; i < employees.length; i++) {
    var employeeName = employees[i];

    // Collect the employee-specific stats
    var stats = {
      name: employeeName,
      monthlyCount: countAppointmentsForEmployeeThisMonth(employeeName), // Monthly count for the employee
      quarterlyCount: countAppointmentsForEmployeeThisQuarter(employeeName), // Quarterly count for the employee
      yearlyCount: countAppointmentsForEmployeeThisYear(employeeName), // Yearly count for the employee
      appointmentsByMonth: getAppointmentsByMonthForEmployee(employeeName)  // **Monthly data for the bar chart**
    };

    // Add the employee stats to the array
    employeeStats.push(stats);
  }

  // Create an HTML template
  var template = HtmlService.createTemplateFromFile('index');
  
  // Pass the total appointment data and employee stats to the HTML template
  template.monthlyCount = monthlyCount;
  template.quarterlyCount = quarterlyCount;
  template.yearlyCount = yearlyCount;
  template.currentQuarter = currentQuarter;
  template.currentMonthName = currentMonthName;
  template.currentYear = currentYear;
  template.appointmentsByMonth = appointmentsByMonth;
  template.employeeStats = JSON.stringify(employeeStats);  // Convert the employee stats array to JSON
  
  // Return the evaluated HTML
  return template.evaluate().setTitle("Tech Connect Metrics (WIP)");
}

