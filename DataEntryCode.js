/**
 * Global variable for the Spreadsheet ID.
 * REPLACE 'YOUR_SPREADSHEET_ID_HERE' with the ID from your Sheet's URL.
 * The ID is the long string between /d/ and /edit.
 */
const SPREADSHEET_ID = '16N9hWBuBlwnA598ci5tKM8252-WORzQNe7TRg8YkHgQ'; 

/**
 * Reusable function to get all non-empty values from the first column (A) of a sheet.
 * The first row (A1) is assumed to be a header and is skipped for robust data loading.
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
    
    // If the sheet is empty or only has a potential header row, return empty data
    // We assume data starts from row 2 (A2), skipping the header in A1.
    if (lastRow <= 1) {
      Logger.log(`Sheet '${sheetName}' has no data entries (only header or empty).`);
      return { data: [] }; 
    }
    
    // Get all values from the first column (A) starting from row 2 (A2) up to the last row with content.
    // Range is (startRow, startCol, numRows, numCols)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 1); 
    const values = dataRange.getValues();
    
    // Filter out empty cells and flatten the 2D array
    const data = values
      .map(row => row[0])
      .filter(cell => cell && cell.toString().trim() !== '');
      
    if (data.length === 0) {
      Logger.log(`No non-empty data found after header in sheet '${sheetName}'.`);
      return { data: [] }; 
    }
    
    return { data: data };

  } catch (e) {
    Logger.log(`Error reading sheet '${sheetName}': ${e.toString()}`);
    return { error: `Failed to load data from '${sheetName}': ${e.toString()}` };
  }
}

/**
 * Reads the 'Schools' sheet and returns a structured object of school names and their available classes.
 * Assumes the first row contains school names (headers) and subsequent rows under each column are the classes for that school.
 * @return {Object} An object mapping school names (string keys) to an array of classes (string array), or an object with an 'error' property.
 */
function getSchoolData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const schoolSheet = ss.getSheetByName('Schools'); 

    if (!schoolSheet) {
      Logger.log("The 'Schools' sheet was not found.");
      return { error: "The 'Schools' worksheet is missing. Please create a sheet named 'Schools'." }; 
    }

    // Get all data from the sheet. Assuming school names are in the first row.
    const dataRange = schoolSheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length < 1 || values[0].every(cell => !cell || cell.toString().trim() === '')) {
      Logger.log("The 'Schools' sheet is empty or only contains empty headers.");
      return {}; // Return empty map
    }

    const headers = values[0]; // First row contains School Names
    const schoolData = {};

    // Iterate through columns (schools) starting from the first header
    for (let col = 0; col < headers.length; col++) {
      const schoolName = headers[col].toString().trim();
      
      if (!schoolName) continue; // Skip empty columns

      // Collect all classes in this column (from row 1 onwards, skipping header row 0)
      const classes = [];
      for (let row = 1; row < values.length; row++) {
        const classValue = values[row][col];
        if (classValue && classValue.toString().trim() !== '') {
          classes.push(classValue.toString().trim());
        }
      }
      
      // We still include the school name even if it has no classes, just in case the front end needs the list of schools
      schoolData[schoolName] = classes;
    }
    
    return schoolData;

  } catch (e) {
    Logger.log('Error reading school data: ' + e.toString());
    return { error: 'Failed to load school data: ' + e.message };
  }
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

/**
 * Reads the 'Schools' sheet and returns a list of school names (headers).
 */
function getSchoolNames() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const schoolSheet = ss.getSheetByName('Schools'); 
    if (!schoolSheet) return [];

    const dataRange = schoolSheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length < 1) return [];

    // School names are in the first row (headers)
    return values[0].filter(name => name && name.toString().trim() !== '');

  } catch (e) {
    Logger.log('Error reading school names: ' + e.toString());
    return [];
  }
}

/**
 * Fetches all necessary data for the filters when the dashboard loads.
 * NEW: Extracts the 'data' array from the getColumnData results for cleaner client consumption.
 */
function getDashboardFilterData() {
  // Use .data or default to empty array []
  const coordinatorNames = getColumnData('Coordinators').data || [];
  const schoolNames = getSchoolNames(); // Already returns a clean array
  const moduleNames = getColumnData('Modules').data || [];

  return {
    coordinators: ['All Coordinators', ...coordinatorNames], // Add the "All" option
    schools: ['All Schools', ...schoolNames],                 // Add the "All" option
    modules: moduleNames
  };
}

// --- Main Data Processing Functions (Already correct) ---

/**
 * Fetches and filters the main session data based on user criteria.
 * @param {Object} filters - Contains name, school, startDate, and endDate.
 * @return {Array} Filtered and processed data rows.
 */
function getFilteredData(filters) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Tracker'); 
  if (!sheet) return [];

  // Get all data, skipping headers
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];

  const headers = values[0];
  let data = values.slice(1);
  
  // Define header indices for easier access and removal
  const indices = {
    name: headers.indexOf('Name'),
    startTime: headers.indexOf('Session Start Time'),
    school: headers.indexOf('School Name'),
    module: headers.indexOf('Module'),
    classCounts: headers.indexOf('Class_Student_Counts'),
    endTime: headers.indexOf('Session End Time'),
    timestamp: headers.indexOf('Timestamp')
  };

  // --- Apply Filters ---
  
  // 1. Coordinator Name Filter
  if (filters.name && filters.name !== 'All Coordinators') {
    data = data.filter(row => row[indices.name] === filters.name);
  }

  // 2. School Name Filter
  if (filters.school && filters.school !== 'All Schools') {
    data = data.filter(row => row[indices.school] === filters.school);
  }

  // 3. Date Range Filter
  if (filters.startDate || filters.endDate) {
    const start = filters.startDate ? new Date(filters.startDate) : null;
    const end = filters.endDate ? new Date(filters.endDate) : null;

    data = data.filter(row => {
      const sessionDate = new Date(row[indices.startTime]);
      
      // Check start date (inclusive)
      const afterStart = start ? sessionDate.getTime() >= start.getTime() : true;
      
      // Check end date (inclusive)
      // Note: We set the time to the end of the day for the end date filter for accurate range coverage
      const beforeEnd = end ? sessionDate.getTime() <= (end.getTime() + 86400000) : true;
      
      return afterStart && beforeEnd;
    });
  }

  // --- Prepare Report 1 (All Sessions) ---
  
  // Determine which columns to KEEP and what their NEW header names should be
  const columnsToKeepIndices = [];
  const report1Headers = ['S No'];

  headers.forEach((header, index) => {
      // Remove 'Timestamp' and 'Session End Time'
      if (index === indices.endTime || index === indices.timestamp) {
          return; 
      }

      // Change 'Session Start Time' header to 'Date'
      if (index === indices.startTime) {
          report1Headers.push('Date');
          columnsToKeepIndices.push(index);
          return;
      }
      
      // Keep all other columns and headers
      report1Headers.push(header);
      columnsToKeepIndices.push(index);
  });
  
  // 4. Create the final data array for the table
  const allSessions = data.map((row, index) => {
    const newRow = [index + 1]; // Start with Serial No.
    
    columnsToKeepIndices.forEach(colIndex => {
      let cellValue = row[colIndex];
      
      // Print only the date part for the Session Start Time (now 'Date')
      if (colIndex === indices.startTime && cellValue instanceof Date) {
        cellValue = cellValue.toLocaleDateString();
      } else if (cellValue instanceof Date) {
        // Format any other remaining Date objects with full locale string
        cellValue = cellValue.toLocaleString(); 
      }
      
      newRow.push(cellValue);
    });
    
    return newRow;
  });


  // 5. Prepare Report 2 (Class Breakdown - Conditional)
  let classBreakdown = null;
  
  if (filters.school && filters.school !== 'All Schools') {
    classBreakdown = {};
    const singleSchoolName = filters.school;
    
    // Get all classes for the selected school (using the School data structure from the other app)
    const schoolSheet = ss.getSheetByName('Schools');
    const schoolValues = schoolSheet.getDataRange().getValues();
    const schoolHeaders = schoolValues[0];
    const schoolColIndex = schoolHeaders.indexOf(singleSchoolName);
    
    if (schoolColIndex !== -1) {
      // Find all classes for the selected school
      const classes = schoolValues.slice(1).map(row => row[schoolColIndex]).filter(c => c);

      // Initialize the breakdown object with all classes
      classes.forEach(cls => {
        classBreakdown[cls] = [];
      });

      // Populate the breakdown with session data
      data.forEach(row => {
        const coordinatorName = row[indices.name]; // Get coordinator name
        const classCountsString = row[indices.classCounts]; // e.g., "6a: 20, 7b: 15"
        const module = row[indices.module];
        // Use the date part for class breakdown consistency
        const startTime = (row[indices.startTime] instanceof Date) ? row[indices.startTime].toLocaleDateString() : row[indices.startTime];
        
        // Parse the class counts string to map class -> student count
        const sessionClassesMap = {};
        if (typeof classCountsString === 'string' && classCountsString.trim() !== '') {
            classCountsString.split(',').forEach(item => {
                const parts = item.split(':').map(s => s.trim());
                if (parts.length === 2 && parts[0] && parts[1]) {
                    const className = parts[0];
                    // Sanitize and convert count, safely handling non-numeric parts
                    const count = parseInt(parts[1].replace(/[^0-9]/g, ''), 10) || 0; 
                    sessionClassesMap[className] = count;
                }
            });
        }
        
        // Iterate over the classes in the current session
        Object.keys(sessionClassesMap).forEach(cls => {
            const studentCount = sessionClassesMap[cls];
            if (classBreakdown.hasOwnProperty(cls)) {
                // Add session details including new coordinator and student count
                classBreakdown[cls].push({
                    module: module,
                    date: startTime,
                    coordinator: coordinatorName,
                    students: studentCount
                });
            }
        });
      });
      
      // Convert the classBreakdown object into an array of objects for easier client-side rendering
      const breakdownArray = Object.keys(classBreakdown).map(cls => ({
        className: cls,
        sessions: classBreakdown[cls]
      }));
      classBreakdown = breakdownArray;

    } else {
        // School exists in filter list but not on Schools sheet structure (shouldn't happen if data is clean)
        classBreakdown = []; 
    }
  }

  return {
    report1: allSessions,
    report1Headers: report1Headers,
    report2: classBreakdown, // null if All Schools selected
    filterValues: filters // return filters used for display
  };
}


/**
 * Handles GET requests to serve the web application's HTML content (Data Entry).
 * This is the entry point for the Data Entry App.
 * @return {HtmlOutput} The HTML content for the web app.
 */
function doGetDataEntry() {
  // Get all data needed for the form
  const schoolData = getSchoolData(); // Returns the map or { error: ... }
  const coordinatorResult = getColumnData('Coordinators');
  const moduleResult = getColumnData('Modules');

  // Create a template object to inject data
  const template = HtmlService.createTemplateFromFile('index');
  
  // UPDATED: Bundle all data, ensuring we pass only the array/map/error message.
  const formData = {
    // Pass the school map (or the error object if present)
    schoolData: schoolData, 
    // Pass the coordinators array (or empty array if error/empty)
    coordinatorData: coordinatorResult.data || [],
    // Pass the modules array (or empty array if error/empty)
    moduleData: moduleResult.data || [],
    // Pass any errors encountered for display on the form
    errors: {
      coordinators: coordinatorResult.error,
      modules: moduleResult.error,
      schools: schoolData.error
    }
  };
  
  // Inject the bundled data into the template
  template.formData = JSON.stringify(formData);

  // Evaluate the template and return the HTML output
  return template.evaluate()
      .setTitle('Chetana Implementation Tracker')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * The default entry point for the DASHBOARD web app.
 */
function doGet() { 
  // Use the new HTML template file
  const template = HtmlService.createTemplateFromFile('DashboardIndex');
  
  // Inject filter options into the HTML
  template.filterData = JSON.stringify(getDashboardFilterData());

  return template.evaluate()
      .setTitle('Session Data Dashboard')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}