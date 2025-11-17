/**
 * Gets all non-empty values from the first column of a specified sheet.
 */
function getColumnData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const dataRange = sheet.getRange(1, 1, sheet.getMaxRows(), 1);
    const values = dataRange.getValues();
    
    return values
      .map(row => row[0])
      .filter(cell => cell && cell.toString().trim() !== '');

  } catch (e) {
    Logger.log(`Error reading sheet '${sheetName}': ${e.toString()}`);
    return [];
  }
}

/**
 * Reads the 'Schools' sheet and returns a list of school names.
 */
function getSchoolNames() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const schoolSheet = ss.getSheetByName('Schools'); 
    if (!schoolSheet) return [];

    const dataRange = schoolSheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length < 1) return [];

    // School names are in the first row
    return values[0].filter(name => name && name.toString().trim() !== '');

  } catch (e) {
    Logger.log('Error reading school names: ' + e.toString());
    return [];
  }
}

/**
 * Fetches all necessary data for the filters when the dashboard loads.
 */
function getDashboardFilterData() {
  const coordinatorNames = getColumnData('Coordinators');
  const schoolNames = getSchoolNames();
  const moduleNames = getColumnData('Modules');

  return {
    coordinators: coordinatorNames,
    schools: schoolNames,
    modules: moduleNames // Not strictly needed for the current report, but good practice
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
 * RENAMED to the default entry point.
 * This is the Web App entry point for the DASHBOARD.
 */
function doGet() { // RENAMED
  // Use the new HTML template file
  const template = HtmlService.createTemplateFromFile('DashboardIndex');
  
  // Inject filter options into the HTML
  template.filterData = JSON.stringify(getDashboardFilterData());

  return template.evaluate()
      .setTitle('Session Data Dashboard')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
