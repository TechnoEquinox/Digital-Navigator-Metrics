// Fetch data from the spreadsheet, analyze the "NEEDS" entries,
// categorizes the entries into call, email, and attention,
// calculate how many days have passed since each entry
function getNeedsMetricsData() {
  const spreadsheetId = '<SPREADSHEET ID GOES HERE>';
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName("Voicemail Log");

  if (!sheet) {
    Logger.log("Sheet 'Voicemail Log' not found.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let needsEntries = [];
  let needsCounts = { call: 0, email: 0, attention: 0 };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const endResult = row[7]; // "End Result" column

    if (endResult && endResult.includes("NEEDS")) {
      if (endResult.includes("CALL")) {
        needsCounts.call++;
        Logger.log(`Row ${i + 1}: Categorized as NEEDS CALL. Total: ${needsCounts.call}`);
      } 
      if (endResult.includes("EMAIL")) {
        needsCounts.email++;
        Logger.log(`Row ${i + 1}: Categorized as NEEDS EMAIL. Total: ${needsCounts.email}`);
      } 
      if (endResult.includes("ATTENTION")) {
        needsCounts.attention++;
        Logger.log(`Row ${i + 1}: Categorized as NEEDS ATTENTION. Total: ${needsCounts.attention}`);
      } 

      // Get details for each entry
      const date = new Date(row[0]);
      date.setHours(0, 0, 0, 0);
      const daysSince = Math.floor((today - date) / (1000 * 60 * 60 * 24));

      needsEntries.push({
        name: row[2],
        phone: row[3],
        reason: row[4],
        archived: row[6],
        daysSince: daysSince
      });
    }
  }

  // Sort by daysSince in descending order
  needsEntries.sort((a, b) => b.daysSince - a.daysSince);

  // Pass data to HTML
  return { needsCounts, needsEntries };
}

// Function to render the HTML page
function doGet() {
  // Start measuring execution time
  var startTime = new Date().getTime();

  const template = HtmlService.createTemplateFromFile("index");
  const metricsData = getNeedsMetricsData();

  // End measuring execution time
  var endTime = new Date().getTime();
  var executionTime = (endTime - startTime) / 1000; // in seconds

  template.metricsData = metricsData;
  template.executionTime = executionTime;
  return template.evaluate().setTitle("Tech Connect Call Log Metrics");
}
