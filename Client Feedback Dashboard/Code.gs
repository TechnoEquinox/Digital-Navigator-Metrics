function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Digital Navigation Feedback Stats');
}

function getFeedbackData() {
  const ss = SpreadsheetApp.openById("<INSERT SPREADSHEET ID HERE>");
  const sheet = ss.getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return {
      ratingCounts: {1: 0, 2: 0, 3: 0, 4: 0, 5: 0},
      referralCounts: {},
      averageRating: "N/A" // Ensure this is always included
    };
  }

  const headers = data[0];
  const ratingColumnName = "How would you rate this session overall? (1 is Poor, 5 is Excellent)";
  const referralColumnName = "How did you hear about our Digital Navigation program?";

  const ratingColIndex = headers.indexOf(ratingColumnName);
  const referralColIndex = headers.indexOf(referralColumnName);

  if (ratingColIndex === -1 || referralColIndex === -1) {
    throw new Error(`One or more columns not found in the sheet`);
  }

  let ratingCounts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0};
  let referralCounts = {};
  let totalRating = 0;
  let totalResponses = 0;

  for (let i = 1; i < data.length; i++) {
    const rating = parseInt(data[i][ratingColIndex]);
    const referral = data[i][referralColIndex]?.trim();

    if (!isNaN(rating) && rating >= 1 && rating <= 5) {
      ratingCounts[rating]++;
      totalRating += rating;
      totalResponses++;
    }

    if (referral) {
      referralCounts[referral] = (referralCounts[referral] || 0) + 1;
    }
  }

  const averageRating = totalResponses > 0 ? (totalRating / totalResponses).toFixed(2) : "N/A";

  return {
    ratingCounts,
    referralCounts,
    averageRating
  };
}

function getNavigatorData() {
  const employees = ["Connor Bailey", "Elijah Mitchell"];

  const ss = SpreadsheetApp.openById("1zGGMrf2uFslvqIh23zGxpa5eAkTc8tu1u-mScwQqDag");
  const sheet = ss.getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return employees.map(name => ({
      name,
      ratingCounts: {1: 0, 2: 0, 3: 0, 4: 0, 5: 0},
      averageRating: "N/A"
    }));
  }

  const ratingMap = {
    "Excellent": 5,
    "Very Good": 4,
    "Good": 3,
    "Fair": 2,
    "Poor": 1
  };

  const headers = data[0];
  const ratingColumnName = "How would you rate this session overall? (1 is Poor, 5 is Excellent)";
  const employeeColumnName = "Which Digital Navigator did you work with?";
  const techKnowledgeColumnName = "Please rate your Navigator's technical knowledge.";
  const communicationColumnName = "Please rate your Navigator's communication skills.";
  const comfortColumnName = "How comfortable did you feel with technology after this appointment?";
  const recommendColumnName = "Would you recommend our Digital Navigation program to others?";
  const commentColumnName = "Additional feedback for your Digital Navigator";
  const timestampColumnName = "Timestamp";

  const ratingColIndex = headers.indexOf(ratingColumnName);
  const employeeColIndex = headers.indexOf(employeeColumnName);
  const techKnowledgeColIndex = headers.indexOf(techKnowledgeColumnName);
  const communicationColIndex = headers.indexOf(communicationColumnName);
  const comfortColIndex = headers.indexOf(comfortColumnName);
  const recommendColIndex = headers.indexOf(recommendColumnName);
  const commentColIndex = headers.indexOf(commentColumnName);
  const timestampColIndex = headers.indexOf(timestampColumnName);

  if (ratingColIndex === -1 || employeeColIndex === -1 || 
    techKnowledgeColIndex === -1 || communicationColIndex === -1 || 
    comfortColIndex === -1 || recommendColIndex === -1 || 
    commentColIndex === -1 || timestampColIndex === -1) {
    
    throw new Error("One or more columns not found in the sheet");
  }

  let employeeStats = employees.map(name => ({
    name,
    ratingCounts: {1: 0, 2: 0, 3: 0, 4: 0, 5: 0},
    totalRating: 0,
    totalResponses: 0,
    totalTechKnowledge: 0,
    techResponses: 0,
    totalCommunication: 0,
    communicationResponses: 0,
    totalComfort: 0,
    comfortResponses: 0,
    recommendYes: 0,
    recommendTotal: 0,
    comments: []
  }));

  for (let i = 1; i < data.length; i++) {
    const rating = parseInt(data[i][ratingColIndex]);
    const navigator = data[i][employeeColIndex]?.trim();
    const techKnowledge = data[i][techKnowledgeColIndex]?.trim();
    const communication = data[i][communicationColIndex]?.trim();
    const comfort = data[i][comfortColIndex]?.trim();
    const recommend = data[i][recommendColIndex]?.trim();

    const navigatorData = employeeStats.find(e => e.name === navigator);
    if (!navigatorData) continue;

    if (!isNaN(rating) && rating >= 1 && rating <= 5) {
      navigatorData.ratingCounts[rating]++;
      navigatorData.totalRating += rating;
      navigatorData.totalResponses++;
    }

    const techScore = ratingMap[techKnowledge];
    if (techScore) {
      navigatorData.totalTechKnowledge += techScore;
      navigatorData.techResponses++;
    }

    const communicationScore = ratingMap[communication];
    if (communicationScore) {
      navigatorData.totalCommunication += communicationScore;
      navigatorData.communicationResponses++;
    }

    const comfortScore = ratingMap[comfort];
    if (comfortScore) {
      navigatorData.totalComfort += comfortScore;
      navigatorData.comfortResponses++;
    }

    if (recommend === "Yes" || recommend === "No") {
      navigatorData.recommendYes += (recommend === "Yes") ? 1 : 0;
      navigatorData.recommendTotal++;
    }

    const comment = data[i][commentColIndex]?.trim();
    const timestamp = data[i][timestampColIndex];

    if (comment) {
      navigatorData.comments.push({
        date: new Date(timestamp).toLocaleDateString(),
        text: comment
      });
    }
  }

  return employeeStats.map(e => ({
    name: e.name,
    totalResponses: e.totalResponses,
    ratingCounts: e.ratingCounts,
    averageRating: e.totalResponses > 0 ? (e.totalRating / e.totalResponses).toFixed(2) : "N/A",
    averageTechKnowledge: e.techResponses > 0 ? (e.totalTechKnowledge / e.techResponses).toFixed(2) : "N/A",
    averageCommunication: e.communicationResponses > 0 ? (e.totalCommunication / e.communicationResponses).toFixed(2) : "N/A",
    averageComfort: e.comfortResponses > 0 ? (e.totalComfort / e.comfortResponses).toFixed(2) : "N/A",
    recommendationPercent: e.recommendTotal > 0 ? `${Math.round((e.recommendYes / e.recommendTotal) * 100)}%` : "N/A",
    comments: e.comments
    .sort((a, b) => new Date(b.date) - new Date(a.date))  // newest first
    .slice(0, 5)
  }));
}



