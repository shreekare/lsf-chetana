/**
 * Global variable for the Spreadsheet ID.
 * REPLACE 'YOUR_SPREADSHEET_ID_HERE' with the ID from your Sheet's URL.
 * The ID is the long string between /d/ and /edit.
 */
const SPREADSHEET_ID = '16N9hWBuBlwnA598ci5tKM8252-WORzQNe7TRg8YkHgQ'; 

/**
 * NEW: Reusable function to get all non-empty values from the first column of a sheet.
 * @param {string} sheetName The name of the worksheet to read.
 * @return {Object} An object containing a 'data' array or an 'error' message.
 */
function getColumnData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet '${sheetName}' was not found.`);
      return { error: `Sheet '${sheetName}' is missing.` };
    }
    
    const lastRow = sheet.getLastRow();
    
    // Check if the sheet actually contains any data rows.
    if (lastRow === 0) {
      Logger.log(`No data found in sheet '${sheetName}'.`);
      return { data: [] }; 
    }
    
    // Get all values from the first column (A) from row 1 up to the last row with content.
    const dataRange = sheet.getRange(1, 1, lastRow, 1);
    const values = dataRange.getValues();
    
    // Filter out empty cells and flatten the 2D array
    const data = values
      .map(row => row[0])
      .filter(cell => cell && cell.toString().trim() !== '');
      
    if (data.length === 0) {
      Logger.log(`No non-empty data found in sheet '${sheetName}'.`);
      return { data: [] }; // Return empty array if no non-empty values are found
    }
    
    return { data: data };

  } catch (e) {
    Logger.log(`Error reading sheet '${sheetName}': ${e.toString()}`);
    // Check for "Sheet is empty" error which can happen if the sheet exists but has no data
    if (e.message.includes("Sheet is empty")) {
        return { data: [] };
    }
    return { error: `Failed to load data from '${sheetName}': ${e.toString()}` };
  }
}

/**
 * Reads the 'Schools' sheet and returns a structured object of school names and their available classes.
 * Assumes the first row contains school names (headers) and subsequent rows under each column are the classes for that school.
 * @return {Object} An object mapping school names to an array of classes.
 */
function getSchoolData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const schoolSheet = ss.getSheetByName('Schools'); 

    if (!schoolSheet) {
      // If the 'Schools' sheet is missing, return an error message to display on the form.
      Logger.log("The 'Schools' sheet was not found.");
      return { error: "The 'Schools' worksheet is missing. Please create a sheet named 'Schools'." }; 
    }

    // Get all data from the sheet. Assuming school names are in the first row.
    const dataRange = schoolSheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length < 1) {
      return {}; // Sheet is empty
    }

    const headers = values[0]; // First row contains School Names
    const schoolData = {};

    // Iterate through columns (schools) starting from the first header
    for (let col = 0; col < headers.length; col++) {
      const schoolName = headers[col].toString().trim();
      
      if (!schoolName) continue; // Skip empty columns

      // Collect all classes in this column (from row 1 onwards)
      const classes = [];
      for (let row = 1; row < values.length; row++) {
        const classValue = values[row][col];
        if (classValue) {
          classes.push(classValue.toString().trim());
        }
      }
      // Only add school if it has classes
      if (classes.length > 0) {
        schoolData[schoolName] = classes;
      }
    }
    
    return schoolData;

  } catch (e) {
    Logger.log('Error reading school data: ' + e.toString());
    // Return a structured error, but still allow the form to load if possible
    return { error: 'Failed to load school data: ' + e.message };
  }
}


/**
 * Handles GET requests to serve the web application's HTML content.
 * This is the entry point when the deployed web app URL is visited.
 * @return {HtmlOutput} The HTML content for the web app.
 */
function doGet() {
  // Get ALL data needed for the form
  const schoolData = getSchoolData();
  const coordinatorData = getColumnData('Coordinators');
  const moduleData = getColumnData('Modules');

  // Create a template object to inject data
  const template = HtmlService.createTemplateFromFile('index');
  
  // UPDATED: Bundle all data into a single object
  const formData = {
    schoolData: schoolData,
    coordinatorData: coordinatorData,
    moduleData: moduleData
  };
  
  // Inject the bundled data into the template
  template.formData = JSON.stringify(formData);

  // Evaluate the template and return the HTML output
  return template.evaluate()
      .setTitle('Chetana Implementation Tracker')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Handles the data submitted from the web form.
 * This function is called by the client-side JavaScript via google.script.run.
 * @param {Object} formData - An object containing the submitted form data.
 */
function processForm(formData) {
  try {
    // 1. Get the target spreadsheet and sheet (Tracker)
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Ensure you target the correct sheet name, default is 'Tracker'
    const sheet = ss.getSheetByName('Tracker'); 

    if (!sheet) {
      throw new Error("Could not find 'Tracker' in the spreadsheet. Please check the sheet name.");
    }

    // 2. Prepare the data array for insertion to match the NEW headers:
    // Timestamp | Name | Session Start Time | Session End Time | Module | School Name | Class_Student_Counts | Learning Outcomes | Challenges Faced
    const rowData = [
      new Date(),                         // Timestamp
      formData.name,                      // Name
      formData.sessionStartTime,          // Session Start Time
      formData.sessionEndTime,            // Session End Time
      formData.module,                    // Module
      formData.schoolName,                // School Name
      formData.classCounts,               // Formatted string (e.g., "6a: 15, 7b: 20")
      formData.learningOutcomes,          // NEW: Learning Outcomes
      formData.challengesFaced            // NEW: Challenges Faced
    ];

    // 3. Append the new row to the sheet
    sheet.appendRow(rowData);

    Logger.log('Data successfully logged to sheet: ' + JSON.stringify(formData));

    return "Data saved successfully";

  } catch (e) {
    Logger.log('Error processing form: ' + e.toString());
    // Rethrow the error so the client-side failure handler can catch it.
    throw new Error('Server error: ' + e.message);
  }
}
