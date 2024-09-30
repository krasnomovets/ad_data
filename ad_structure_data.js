function checkDataStructure() {
  Logger.log("Starting checkDataStructure function");
  const sheetId = ''; // GOOGLE SHEET ID
  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
    Logger.log("Successfully opened spreadsheet with ID: " + sheetId);
  } catch (e) {
    Logger.log("Error opening spreadsheet: " + e.toString());
    sendAlertEmail("Error opening the spreadsheet. Please check the script's permissions and the spreadsheet ID.");
    return false;
  }
  
  const rawData = ss.getSheetByName('Raw Data'); // SHEET NAME "Raw Data"
  
  if (!rawData) {
    Logger.log("Error: Raw Data sheet not found");
    sendAlertEmail("Raw Data sheet is missing in the spreadsheet.");
    return false;
  }

  // Expected column headers and their positions
  const expectedHeaders = {
    'time_start': 0,
    'time_end': 1,
    'model': 2,
    'launch_place': 3,
    'target': 4,
    'carrier': 5,
    'launched': 6,
    'destroyed': 7,
    'not_reach_goal': 8,
    'cross_border_belarus': 9,
    'back_russia': 10,
    'destroyed_details': 11,
    'launched_details': 12,
    'launch_place_details': 13,
    'source': 14
  };

  // Get the actual headers from the Raw Data sheet
  const actualHeaders = rawData.getRange(1, 1, 1, rawData.getLastColumn()).getValues()[0];
  
  let structureChanged = false;
  let changes = [];

  // Check if each expected header is in the correct position
  Object.entries(expectedHeaders).forEach(([header, position]) => {
    if (actualHeaders[position] !== header) {
      structureChanged = true;
      changes.push(`Expected '${header}' in column ${position + 1}, found '${actualHeaders[position] || 'empty'}'`);
    }
  });

  if (structureChanged) {
    const message = "Data structure in Raw Data has changed:\n" + changes.join("\n");
    Logger.log(message);
    sendAlertEmail(message);
    return false;
  }

  Logger.log("Data structure check passed");
  return true;
}

function cleanData() {
  Logger.log("Starting cleanData function");
  
  // Check data structure before proceeding
  if (!checkDataStructure()) {
    Logger.log("Aborting cleanData due to data structure mismatch");
    return;
  }

  const sheetId = ''; // SHEET ID
  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
    Logger.log("Successfully opened spreadsheet for cleaning with ID: " + sheetId);
  } catch (e) {
    Logger.log("Error opening spreadsheet for cleaning: " + e.toString());
    sendAlertEmail("Error opening the spreadsheet for cleaning. Please check the script's permissions and the spreadsheet ID.");
    return;
  }

  const rawData = ss.getSheetByName('Raw Data'); // Sheet name
  const cleanedData = ss.getSheetByName('Cleaned Data') || ss.insertSheet('Cleaned Data'); // Sheet name

  // Clear previous data
  cleanedData.clear();
  cleanedData.appendRow(['Date', 'Model', 'Type', 'Target', 'Launched', 'Destroyed']);

  // Get data from Raw Data sheet
  const data = rawData.getRange(2, 1, rawData.getLastRow() - 1, 15).getValues(); 
  Logger.log(`Retrieved ${data.length} rows from Raw Data sheet`);

  // List of models to filter out
  const modelsToFilter = ['Zala', 'Supercam', 'Orlan-10', 'Unknown Missile', 'Reconnaissance UAV', 'ZALA', 'Merlin-VR', 
                          'Unknown UAV', 'Forpost', 'Lancet', 'KAB', 'Mohajer-6', 'Orlan-10 and ZALA and Supercam',
                          'Orlan-10 and ZALA', 'Orlan-10 and Orlan-30 and ZALA and Supercam', 'Orlan-30', 
                          'Orlan-10 and Supercam', 'Granat-4', 'Orion','Картограф', 'Привет-82'];

  // Initialize unclassified models log
  let unclassifiedModels = new Set();

  // Process each row from Raw Data
  const cleanedValues = data.filter(row => {
    const date = new Date(row[0]); // Using time_start as the date
    const month = date.getMonth() + 1;
    const year = date.getFullYear();
    return !modelsToFilter.includes(row[2]) && !(year === 2022 && month === 9);
  }).map((row, index) => {
    const date = formatDate(new Date(row[0])); // Using time_start as the date
    const model = row[2];
    const type = classifyModel(model);
    const target = row[4];
    const launched = parseInt(row[6]);
    const destroyed = parseInt(row[7]);

    if (type === 'Unclassified') {
      unclassifiedModels.add(model);
    }

    // Log every 100th row for debugging
    if (index % 100 === 0) {
      Logger.log(`Processing row ${index}: Date=${date}, Model=${model}, Type=${type}, Target=${target}, Launched=${launched}, Destroyed=${destroyed}`);
    }

    // Check for zero values
    if (launched === 0 || destroyed === 0) {
      Logger.log(`Warning: Zero value detected in row ${index + 2} of Raw Data. Model: ${model}, Launched: ${launched}, Destroyed: ${destroyed}`);
    }

    return [date, model, type, target, launched, destroyed];
  });

  Logger.log(`Processed ${cleanedValues.length} rows after filtering`);

  // Write the cleaned and classified data back to the Cleaned Data sheet
  if (cleanedValues.length > 0) {
    cleanedData.getRange(2, 1, cleanedValues.length, 6).setValues(cleanedValues);
    Logger.log(`Wrote ${cleanedValues.length} rows to Cleaned Data sheet`);
  } else {
    Logger.log("No data to write to Cleaned Data sheet");
  }

  // Send email if there are unclassified models
  if (unclassifiedModels.size > 0) {
    Logger.log(`Unclassified models found: ${Array.from(unclassifiedModels).join(', ')}`);
    sendAlertEmail("Unclassified Models Detected", 
                   "There was an unclassified missile model in Clean Air Def Data script. The following models could not be classified and require your attention: " + Array.from(unclassifiedModels).join(', '));
  } else {
    Logger.log("No unclassified models found");
  }
}

function sendAlertEmail(subject, message) {
  MailApp.sendEmail({
    to: "qpkrasnomovets@gmail.com",
    subject: subject,
    body: message
  });
}


function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
}

function classifyModel(model) {
  const rules = [
    { keyword: ['C-300', 'C-400', 'Iskander', 'KN-23'], type: 'ballistic' },
    { keyword: ['X-101', 'X-555', 'X-35', 'X-59', 'X-69', 'Kalibr'], type: 'cruise subsonic missile' },
    { keyword: ['P-800', 'X-22', 'X-32', 'X-31'], type: 'cruise supersonic missile' },
    { keyword: ['Zircon', 'Kinzhal', 'X-47'], type: 'hypersonic' },
    { keyword: ['Shahed'], type: 'shahed drone' }
  ];

  for (const rule of rules) {
    if (rule.keyword.some(kw => model.includes(kw))) {
      return rule.type;
    }
  }
  return 'Unclassified'; // Return this if no rules match
}