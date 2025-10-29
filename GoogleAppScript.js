const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const MASTER_LOG_SHEET_NAME = 'Master Job Log';
const SCHEDULE_SHEET_NAME = 'Onsite Schedule';
const CALENDAR_NAME = 'Pacific NW Computers';

function doPost(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  
    // Check if this is an update action
  if (e.parameter && e.parameter.action === 'updateJob') {
    return updateJob(e);
  }

  if (!sheet) {
    return createJsonResponse({ result: 'error', message: 'Sheet not found' });
  }
  
  // Handle form data sent via URLSearchParams
  let data;
  try {
    // The data comes in as e.parameter.data (already a string)
    if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);  // Parse the JSON string
    } else if (e.postData && e.postData.contents) {
      // Fallback for old JSON format
      data = JSON.parse(e.postData.contents);
    } else {
      throw new Error('No data received');
    }
  } catch (error) {
    Logger.log('Parse error: ' + error.message);
    Logger.log('Received parameter: ' + JSON.stringify(e.parameter));
    return createJsonResponse({ result: 'error', message: 'Invalid data format: ' + error.message });
  }

  const rowData = [
    data.jobId || '',
    new Date(),
    data.serviceType || '',
    data.clientName || '',
    data.clientEmail || '',
    data.clientPhone || '',
    data.systemMake || '',
    data.dueDate ? new Date(data.dueDate) : '',
    data.currentStatus || '1. Checked In: Diagnostics',
    'Technician Name',
    data.initialRequest || '',
    '',
    '',
    ''
  ];

  try {
    sheet.appendRow(rowData);
    SpreadsheetApp.flush();
    return createJsonResponse({ result: 'success', message: `Job ${data.jobId} logged successfully!` });
  } catch (error) {
    Logger.log('Sheet append error: ' + error.message);
    return createJsonResponse({ result: 'error', message: error.message });
  }
}

function doGet(e) {
  const action = e.parameter.action;
  const status = e.parameter.status;
  const sheetParam = e.parameter.sheet;
  const rowIndex = e.parameter.rowIndex;  // ← ADD THIS - was missing!
  
  // NEW: Handle get all jobs request (for Job Log view)
  if (action === 'getAllJobs') {
    return getAllJobs();
  }
  
  // NEW: Handle get single job by row index (for editing)
  if (action === 'getJobByRow') {
    return getJobByRow(rowIndex);
  }
  
  // Handle job details request by status (for clickable status sections)
  if (action === 'getJobsByStatus') {
    return getJobsByStatus(status);
  }
  
  // Handle status counts request (for dashboard status counts)
  if (sheetParam === 'MasterLogStatus') {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
    const statusCounts = getMasterLogStatusCounts(sheet);
    return createJsonResponse(statusCounts);
  }
  
  // Default: Return schedule data (for dashboard schedule section)
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const scheduleData = getSheetDataAsJson(scheduleSheet);
  return createJsonResponse(scheduleData);
}

// NEW FUNCTION: Get jobs by status
function getJobsByStatus(status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master Job Log'); // Use your actual sheet name
  const data = sheet.getDataRange().getValues();
  
  // Assuming first row is headers
  const headers = data[0];
  const jobs = [];
  
  // Find the column index for Status
  const statusColIndex = headers.indexOf('Status'); // Adjust column name if different
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const jobStatus = row[statusColIndex];
    
    // If status is 'all', include all jobs, otherwise filter by status
    if (status === 'all' || jobStatus === status) {
      // Create job object with all relevant fields
      const job = {
      Job_ID: row[headers.indexOf('Job ID')],
      Client_Name: row[headers.indexOf('Client Name')],
      Client_Email: row[headers.indexOf('Client Email')],
      Client_Phone: row[headers.indexOf('Client Phone')],
      Service_Type: row[headers.indexOf('Service Type')],
      Due_Date: row[headers.indexOf('Due Date')] ? Utilities.formatDate(new Date(row[headers.indexOf('Due Date')]), Session.getScriptTimeZone(), 'MM/dd/yyyy') : '',
      Status: jobStatus,
      System_Make_Model: row[headers.indexOf('System Make/Model')],  // ← CHANGED: / instead of &
      Initial_Request: row[headers.indexOf('Initial Request')]
    };
      jobs.push(job);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(jobs))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetDataAsJson(sheet) {
  if (!sheet) throw new Error('Sheet not found');
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).replace(/[\s\/]/g, '_'));
  const data = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = values[i][j];
    }
    data.push(row);
  }
  return data;
}

function getMasterLogStatusCounts(sheet) {
  if (!sheet) throw new Error('Sheet not found');
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  if (values.length < 2) {
    return {
      '1. Checked In: Diagnostics': 0,
      '2. Awaiting Customer Approval': 0,
      '3. Awaiting Parts (Vendor Side)': 0,
      '4. In Progress: Repair/Install': 0,
      '5. Ready for Pickup/Delivery': 0
    };
  }
  
  // Find Status column dynamically
  const headers = values[0];
  const statusColIndex = headers.indexOf('Status');
  
  if (statusColIndex === -1) {
    throw new Error('Status column not found in sheet');
  }
  
  const counts = {
    '1. Checked In: Diagnostics': 0,
    '2. Awaiting Customer Approval': 0,
    '3. Awaiting Parts (Vendor Side)': 0,
    '4. In Progress: Repair/Install': 0,
    '5. Ready for Pickup/Delivery': 0
  };
  
  for (let i = 1; i < values.length; i++) {
    const status = values[i][statusColIndex];
    if (counts.hasOwnProperty(status)) {
      counts[status]++;
    }
  }
  
  return counts;
}

function syncOnsiteSchedule() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME);

  if (!calendar || calendar.length === 0) {
    Logger.log('Calendar not found: ' + CALENDAR_NAME);
    return;
  }
  
  const calendarObj = calendar[0];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const twoWeeks = new Date();
  twoWeeks.setDate(today.getDate() + 14);
  twoWeeks.setHours(23, 59, 59, 999);

  const events = calendarObj.getEvents(today, twoWeeks);
  const newScheduleData = [];

  if (scheduleSheet.getLastRow() > 1) {
    scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow() - 1, scheduleSheet.getLastColumn()).clearContent();
  }

  for (const event of events) {
    if (!event.isAllDayEvent()) {
      const eventTitle = event.getTitle();
      const eventDesc = event.getDescription();
      const jobIdMatch = eventTitle.match(/WO-\d{4}/) || eventDesc.match(/WO-\d{4}/);
      const jobId = jobIdMatch ? jobIdMatch[0] : '';
      
      const newRow = [
        event.getStartTime(),
        `${Utilities.formatDate(event.getStartTime(), ss.getSpreadsheetTimeZone(), 'hh:mm a')} - ${Utilities.formatDate(event.getEndTime(), ss.getSpreadsheetTimeZone(), 'hh:mm a')}`,
        jobId,
        eventTitle,
        event.getLocation() || '',
        eventDesc || ''
      ];
      newScheduleData.push(newRow);
    }
  }
  
  if (newScheduleData.length > 0) {
    scheduleSheet.getRange(2, 1, newScheduleData.length, newScheduleData[0].length).setValues(newScheduleData);
  }
  
  SpreadsheetApp.flush();
  Logger.log('Synced ' + newScheduleData.length + ' events');
}

function createJsonResponse(obj, status = 200) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function setupJobTrackingSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let masterLogSheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!masterLogSheet) {
    masterLogSheet = ss.insertSheet(MASTER_LOG_SHEET_NAME, 0);
    const masterHeaders = [
      'Job ID', 'Date In', 'Service Type', 'Client Name', 'Client Email', 'Client Phone', 
      'System Make/Model', 'Due Date', 'Status', 'Technician', 'Initial Request', 
      'Job Notes', 'Final Resolution', 'Date Completed'
    ];
    masterLogSheet.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders]).setFontWeight('bold');
  }

  let scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet(SCHEDULE_SHEET_NAME);
    const scheduleHeaders = ['Event Date', 'Time Start/End', 'Job ID', 'Client Name', 'Location/Address', 'Event Notes'];
    scheduleSheet.getRange(1, 1, 1, scheduleHeaders.length).setValues([scheduleHeaders]).setFontWeight('bold');
  }

  ss.setActiveSheet(masterLogSheet);
  Logger.log('Setup complete');
}

function getAllJobs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    return createJsonResponse([]);
  }
  
  const headers = data[0];
  const jobs = [];
  
  // Find column indices
  const jobIdCol = headers.indexOf('Job ID');
  const dateInCol = headers.indexOf('Date In');
  const serviceTypeCol = headers.indexOf('Service Type');
  const clientNameCol = headers.indexOf('Client Name');
  const clientEmailCol = headers.indexOf('Client Email');
  const clientPhoneCol = headers.indexOf('Client Phone');
  const systemCol = headers.indexOf('System Make/Model');
  const dueDateCol = headers.indexOf('Due Date');
  const statusCol = headers.indexOf('Status');
  const technicianCol = headers.indexOf('Technician');
  const requestCol = headers.indexOf('Initial Request');
  const notesCol = headers.indexOf('Job Notes');
  const resolutionCol = headers.indexOf('Final Resolution');
  const completedCol = headers.indexOf('Date Completed');
  
  // Process each row (skip header row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const job = {
      rowIndex: i + 1, // Row index in sheet (1-based, accounting for header)
      Job_ID: row[jobIdCol] || '',
      Date_In: row[dateInCol] ? formatDate(row[dateInCol]) : '',
      Service_Type: row[serviceTypeCol] || '',
      Client_Name: row[clientNameCol] || '',
      Client_Email: row[clientEmailCol] || '',
      Client_Phone: row[clientPhoneCol] || '',
      System_Make_Model: row[systemCol] || '',
      Due_Date: row[dueDateCol] ? formatDate(row[dueDateCol]) : '',
      Status: row[statusCol] || '',
      Technician: row[technicianCol] || '',
      Initial_Request: row[requestCol] || '',
      Job_Notes: row[notesCol] || '',
      Final_Resolution: row[resolutionCol] || '',
      Date_Completed: row[completedCol] ? formatDate(row[completedCol]) : ''
    };
    jobs.push(job);
  }
  
  return createJsonResponse(jobs);
}

// NEW FUNCTION: Get a single job by row index
function getJobByRow(rowIndex) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  if (!rowIndex || rowIndex < 2 || rowIndex > data.length) {
    return createJsonResponse({ result: 'error', message: 'Invalid row index' });
  }
  
  const headers = data[0];
  const row = data[rowIndex - 1]; // Convert to 0-based index
  
  // Find column indices
  const jobIdCol = headers.indexOf('Job ID');
  const dateInCol = headers.indexOf('Date In');
  const serviceTypeCol = headers.indexOf('Service Type');
  const clientNameCol = headers.indexOf('Client Name');
  const clientEmailCol = headers.indexOf('Client Email');
  const clientPhoneCol = headers.indexOf('Client Phone');
  const systemCol = headers.indexOf('System Make/Model');
  const dueDateCol = headers.indexOf('Due Date');
  const statusCol = headers.indexOf('Status');
  const technicianCol = headers.indexOf('Technician');
  const requestCol = headers.indexOf('Initial Request');
  const notesCol = headers.indexOf('Job Notes');
  const resolutionCol = headers.indexOf('Final Resolution');
  const completedCol = headers.indexOf('Date Completed');
  
  const job = {
    rowIndex: rowIndex,
    Job_ID: row[jobIdCol] || '',
    Date_In: row[dateInCol] ? formatDate(row[dateInCol]) : '',
    Service_Type: row[serviceTypeCol] || '',
    Client_Name: row[clientNameCol] || '',
    Client_Email: row[clientEmailCol] || '',
    Client_Phone: row[clientPhoneCol] || '',
    System_Make_Model: row[systemCol] || '',
    Due_Date: row[dueDateCol] ? formatDateForInput(row[dueDateCol]) : '',
    Status: row[statusCol] || '',
    Technician: row[technicianCol] || '',
    Initial_Request: row[requestCol] || '',
    Job_Notes: row[notesCol] || '',
    Final_Resolution: row[resolutionCol] || '',
    Date_Completed: row[completedCol] ? formatDateForInput(row[completedCol]) : ''
  };
  
  return createJsonResponse(job);
}

// NEW FUNCTION: Update a job in the Master Job Log
function updateJob(e) {
  try {
    const data = JSON.parse(e.parameter.data);
    const rowIndex = data.rowIndex;
    
    if (!rowIndex || rowIndex < 2) {
      return createJsonResponse({ result: 'error', message: 'Invalid row index' });
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find column indices
    const serviceTypeCol = headers.indexOf('Service Type') + 1;
    const clientNameCol = headers.indexOf('Client Name') + 1;
    const clientEmailCol = headers.indexOf('Client Email') + 1;
    const clientPhoneCol = headers.indexOf('Client Phone') + 1;
    const systemCol = headers.indexOf('System Make/Model') + 1;
    const dueDateCol = headers.indexOf('Due Date') + 1;
    const statusCol = headers.indexOf('Status') + 1;
    const requestCol = headers.indexOf('Initial Request') + 1;
    const notesCol = headers.indexOf('Job Notes') + 1;
    
    // Update the cells
    if (clientNameCol > 0) sheet.getRange(rowIndex, clientNameCol).setValue(data.Client_Name || '');
    if (clientEmailCol > 0) sheet.getRange(rowIndex, clientEmailCol).setValue(data.Client_Email || '');
    if (clientPhoneCol > 0) sheet.getRange(rowIndex, clientPhoneCol).setValue(data.Client_Phone || '');
    if (serviceTypeCol > 0) sheet.getRange(rowIndex, serviceTypeCol).setValue(data.Service_Type || '');
    if (dueDateCol > 0) sheet.getRange(rowIndex, dueDateCol).setValue(data.Due_Date ? new Date(data.Due_Date) : '');
    if (statusCol > 0) sheet.getRange(rowIndex, statusCol).setValue(data.Status || '');
    if (systemCol > 0) sheet.getRange(rowIndex, systemCol).setValue(data.System_Make_Model || '');
    if (requestCol > 0) sheet.getRange(rowIndex, requestCol).setValue(data.Initial_Request || '');
    if (notesCol > 0) sheet.getRange(rowIndex, notesCol).setValue(data.Job_Notes || '');
    
    SpreadsheetApp.flush();
    
    return createJsonResponse({ result: 'success', message: 'Job updated successfully!' });
  } catch (error) {
    Logger.log('Update error: ' + error.message);
    return createJsonResponse({ result: 'error', message: error.message });
  }
}

// Helper function to format dates for display
function formatDate(date) {
  if (!date) return '';
  try {
    if (typeof date === 'string') return date;
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MM/dd/yyyy');
  } catch (e) {
    return date.toString();
  }
}

// Helper function to format dates for input fields (YYYY-MM-DD)
function formatDateForInput(date) {
  if (!date) return '';
  try {
    if (typeof date === 'string' && date.match(/^\d{4}-\d{2}-\d{2}$/)) return date;
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}
