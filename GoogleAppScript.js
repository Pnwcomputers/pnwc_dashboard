const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const MASTER_LOG_SHEET_NAME = 'Master Job Log';
const SCHEDULE_SHEET_NAME = 'Onsite Schedule';
const CALENDAR_NAME = 'Pacific NW Computers';

// ============================================
// EMAIL FUNCTIONS (NEW)
// ============================================

/**
 * Send confirmation email when a new job is checked in
 */
function sendConfirmationEmail(jobData) {
  try {
    const subject = `Service Confirmation - Job #${jobData.jobId} - Pacific NW Computers`;
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #2563eb 0%, #4f46e5 100%); color: white; padding: 30px; text-align: center;">
          <h1 style="margin: 0; font-size: 28px;">Pacific NW Computers</h1>
          <p style="margin: 10px 0 0 0; font-size: 16px;">Service Request Confirmed</p>
        </div>
        
        <div style="padding: 30px; background-color: #f9fafb; border: 1px solid #e5e7eb;">
          <p style="font-size: 16px; color: #374151; margin-top: 0;">Dear ${jobData.clientName},</p>
          
          <p style="font-size: 16px; color: #374151;">Thank you for choosing Pacific NW Computers! Your service request has been received and logged into our system.</p>
          
          <div style="background-color: white; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #2563eb;">
            <h2 style="color: #1f2937; margin-top: 0; font-size: 20px;">Job Details</h2>
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">Job ID:</td>
                <td style="padding: 8px 0; color: #1f2937; font-weight: bold;">${jobData.jobId}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">Service Type:</td>
                <td style="padding: 8px 0; color: #1f2937;">${jobData.serviceType}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">Due Date:</td>
                <td style="padding: 8px 0; color: #1f2937;">${jobData.dueDate || 'TBD'}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">Current Status:</td>
                <td style="padding: 8px 0; color: #1f2937;">${jobData.currentStatus}</td>
              </tr>
              ${jobData.systemMake ? `
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">System:</td>
                <td style="padding: 8px 0; color: #1f2937;">${jobData.systemMake}</td>
              </tr>
              ` : ''}
            </table>
          </div>
          
          ${jobData.initialRequest ? `
          <div style="background-color: #eff6ff; padding: 15px; border-radius: 8px; margin: 20px 0;">
            <h3 style="color: #1e40af; margin-top: 0; font-size: 16px;">Your Request:</h3>
            <p style="color: #374151; margin-bottom: 0;">${jobData.initialRequest}</p>
          </div>
          ` : ''}
          
          <p style="font-size: 16px; color: #374151;">We will keep you updated on the progress of your service. If you have any questions, please don't hesitate to contact us.</p>
          
          <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e5e7eb;">
            <p style="color: #6b7280; font-size: 14px; margin: 5px 0;">Best regards,</p>
            <p style="color: #1f2937; font-size: 16px; font-weight: 600; margin: 5px 0;">Pacific NW Computers Team</p>
          </div>
        </div>
        
        <div style="background-color: #1f2937; color: #9ca3af; padding: 20px; text-align: center; font-size: 12px;">
          <p style="margin: 5px 0;">This is an automated confirmation email.</p>
          <p style="margin: 5px 0;">Please keep this email for your records.</p>
        </div>
      </div>
    `;
    
    const plainBody = `
Dear ${jobData.clientName},

Thank you for choosing Pacific NW Computers! Your service request has been received and logged into our system.

JOB DETAILS
-----------
Job ID: ${jobData.jobId}
Service Type: ${jobData.serviceType}
Due Date: ${jobData.dueDate || 'TBD'}
Current Status: ${jobData.currentStatus}
${jobData.systemMake ? 'System: ' + jobData.systemMake : ''}

${jobData.initialRequest ? 'YOUR REQUEST:\n' + jobData.initialRequest : ''}

We will keep you updated on the progress of your service. If you have any questions, please don't hesitate to contact us.

Best regards,
Pacific NW Computers Team

---
This is an automated confirmation email.
Please keep this email for your records.
    `;
    
    MailApp.sendEmail({
      to: jobData.clientEmail,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });
    
    Logger.log('Confirmation email sent to: ' + jobData.clientEmail);
  } catch (error) {
    Logger.log('Error sending confirmation email: ' + error.message);
  }
}

/**
 * Send update email when job status or notes change
 */
function sendUpdateEmail(jobData, changes) {
  try {
    const subject = `Job Update - #${jobData.Job_ID} - Pacific NW Computers`;
    
    let changesHtml = '';
    if (changes.statusChanged) {
      changesHtml += `
        <div style="background-color: #fef3c7; padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid #f59e0b;">
          <h3 style="color: #92400e; margin-top: 0; font-size: 16px;">✓ Status Updated</h3>
          <p style="color: #78350f; margin-bottom: 0; font-size: 15px;"><strong>${jobData.Status}</strong></p>
        </div>
      `;
    }
    
    if (changes.notesChanged && jobData.Job_Notes) {
      changesHtml += `
        <div style="background-color: #dbeafe; padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid #3b82f6;">
          <h3 style="color: #1e40af; margin-top: 0; font-size: 16px;">📝 New Notes Added</h3>
          <p style="color: #1e3a8a; margin-bottom: 0; white-space: pre-wrap;">${jobData.Job_Notes}</p>
        </div>
      `;
    }
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #2563eb 0%, #4f46e5 100%); color: white; padding: 30px; text-align: center;">
          <h1 style="margin: 0; font-size: 28px;">Pacific NW Computers</h1>
          <p style="margin: 10px 0 0 0; font-size: 16px;">Service Update</p>
        </div>
        
        <div style="padding: 30px; background-color: #f9fafb; border: 1px solid #e5e7eb;">
          <p style="font-size: 16px; color: #374151; margin-top: 0;">Dear ${jobData.Client_Name},</p>
          
          <p style="font-size: 16px; color: #374151;">There's an update on your service request <strong>Job #${jobData.Job_ID}</strong>.</p>
          
          ${changesHtml}
          
          <div style="background-color: white; padding: 20px; border-radius: 8px; margin: 20px 0;">
            <h3 style="color: #1f2937; margin-top: 0; font-size: 18px;">Current Job Information</h3>
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">Job ID:</td>
                <td style="padding: 8px 0; color: #1f2937; font-weight: bold;">${jobData.Job_ID}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">Status:</td>
                <td style="padding: 8px 0; color: #1f2937;">${jobData.Status}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">Service Type:</td>
                <td style="padding: 8px 0; color: #1f2937;">${jobData.Service_Type}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #6b7280; font-weight: 600;">Due Date:</td>
                <td style="padding: 8px 0; color: #1f2937;">${jobData.Due_Date || 'TBD'}</td>
              </tr>
            </table>
          </div>
          
          <p style="font-size: 16px; color: #374151;">If you have any questions about this update, please don't hesitate to contact us.</p>
          
          <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e5e7eb;">
            <p style="color: #6b7280; font-size: 14px; margin: 5px 0;">Best regards,</p>
            <p style="color: #1f2937; font-size: 16px; font-weight: 600; margin: 5px 0;">Pacific NW Computers Team</p>
          </div>
        </div>
        
        <div style="background-color: #1f2937; color: #9ca3af; padding: 20px; text-align: center; font-size: 12px;">
          <p style="margin: 5px 0;">This is an automated notification email.</p>
        </div>
      </div>
    `;
    
    const plainBody = `
Dear ${jobData.Client_Name},

There's an update on your service request Job #${jobData.Job_ID}.

${changes.statusChanged ? `STATUS UPDATED: ${jobData.Status}\n` : ''}
${changes.notesChanged && jobData.Job_Notes ? `\nNEW NOTES:\n${jobData.Job_Notes}\n` : ''}

CURRENT JOB INFORMATION
-----------------------
Job ID: ${jobData.Job_ID}
Status: ${jobData.Status}
Service Type: ${jobData.Service_Type}
Due Date: ${jobData.Due_Date || 'TBD'}

If you have any questions about this update, please don't hesitate to contact us.

Best regards,
Pacific NW Computers Team

---
This is an automated notification email.
    `;
    
    MailApp.sendEmail({
      to: jobData.Client_Email,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });
    
    Logger.log('Update email sent to: ' + jobData.Client_Email);
  } catch (error) {
    Logger.log('Error sending update email: ' + error.message);
  }
}

// ============================================
// ORIGINAL FUNCTIONS (MODIFIED)
// ============================================

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
    
    // ✨ NEW: Send confirmation email
    if (data.clientEmail) {
      sendConfirmationEmail(data);
    }
    
    return createJsonResponse({ result: 'success', message: `Job ${data.jobId} logged successfully! Confirmation email sent.` });
  } catch (error) {
    Logger.log('Sheet append error: ' + error.message);
    return createJsonResponse({ result: 'error', message: error.message });
  }
}

function doGet(e) {
  const action = e.parameter.action;
  const status = e.parameter.status;
  const sheetParam = e.parameter.sheet;
  const rowIndex = e.parameter.rowIndex;
  
  // Handle get all jobs request (for Job Log view)
  if (action === 'getAllJobs') {
    return getAllJobs();
  }
  
  // Handle get single job by row index (for editing)
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

function getJobsByStatus(status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master Job Log');
  const data = sheet.getDataRange().getValues();
  
  const headers = data[0];
  const jobs = [];
  
  const statusColIndex = headers.indexOf('Status');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const jobStatus = row[statusColIndex];
    
    if (status === 'all' || jobStatus === status) {
      const job = {
        Job_ID: row[headers.indexOf('Job ID')],
        Client_Name: row[headers.indexOf('Client Name')],
        Client_Email: row[headers.indexOf('Client Email')],
        Client_Phone: row[headers.indexOf('Client Phone')],
        Service_Type: row[headers.indexOf('Service Type')],
        Due_Date: row[headers.indexOf('Due Date')] ? Utilities.formatDate(new Date(row[headers.indexOf('Due Date')]), Session.getScriptTimeZone(), 'MM/dd/yyyy') : '',
        Status: jobStatus,
        System_Make_Model: row[headers.indexOf('System Make/Model')],
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
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const job = {
      rowIndex: i + 1,
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

function getJobByRow(rowIndex) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  if (!rowIndex || rowIndex < 2 || rowIndex > data.length) {
    return createJsonResponse({ result: 'error', message: 'Invalid row index' });
  }
  
  const headers = data[0];
  const row = data[rowIndex - 1];
  
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

// ✨ MODIFIED: Update job with email notifications
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
    
    // ✨ NEW: Get old values to compare changes
    const oldRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol = headers.indexOf('Status');
    const notesCol = headers.indexOf('Job Notes');
    const oldStatus = oldRow[statusCol];
    const oldNotes = oldRow[notesCol];
    
    // Find column indices
    const serviceTypeCol = headers.indexOf('Service Type') + 1;
    const clientNameCol = headers.indexOf('Client Name') + 1;
    const clientEmailCol = headers.indexOf('Client Email') + 1;
    const clientPhoneCol = headers.indexOf('Client Phone') + 1;
    const systemCol = headers.indexOf('System Make/Model') + 1;
    const dueDateCol = headers.indexOf('Due Date') + 1;
    const statusColNum = headers.indexOf('Status') + 1;
    const requestCol = headers.indexOf('Initial Request') + 1;
    const notesColNum = headers.indexOf('Job Notes') + 1;
    
    // Update the cells
    if (clientNameCol > 0) sheet.getRange(rowIndex, clientNameCol).setValue(data.Client_Name || '');
    if (clientEmailCol > 0) sheet.getRange(rowIndex, clientEmailCol).setValue(data.Client_Email || '');
    if (clientPhoneCol > 0) sheet.getRange(rowIndex, clientPhoneCol).setValue(data.Client_Phone || '');
    if (serviceTypeCol > 0) sheet.getRange(rowIndex, serviceTypeCol).setValue(data.Service_Type || '');
    if (dueDateCol > 0) sheet.getRange(rowIndex, dueDateCol).setValue(data.Due_Date ? new Date(data.Due_Date) : '');
    if (statusColNum > 0) sheet.getRange(rowIndex, statusColNum).setValue(data.Status || '');
    if (systemCol > 0) sheet.getRange(rowIndex, systemCol).setValue(data.System_Make_Model || '');
    if (requestCol > 0) sheet.getRange(rowIndex, requestCol).setValue(data.Initial_Request || '');
    if (notesColNum > 0) sheet.getRange(rowIndex, notesColNum).setValue(data.Job_Notes || '');
    
    SpreadsheetApp.flush();
    
    // ✨ NEW: Check if status or notes changed and send email
    const statusChanged = oldStatus !== data.Status;
    const notesChanged = oldNotes !== data.Job_Notes;
    
    if ((statusChanged || notesChanged) && data.Client_Email) {
      sendUpdateEmail(data, {
        statusChanged: statusChanged,
        notesChanged: notesChanged
      });
    }
    
    return createJsonResponse({ result: 'success', message: 'Job updated successfully!' });
  } catch (error) {
    Logger.log('Update error: ' + error.message);
    return createJsonResponse({ result: 'error', message: error.message });
  }
}

function formatDate(date) {
  if (!date) return '';
  try {
    if (typeof date === 'string') return date;
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MM/dd/yyyy');
  } catch (e) {
    return date.toString();
  }
}

function formatDateForInput(date) {
  if (!date) return '';
  try {
    if (typeof date === 'string' && date.match(/^\d{4}-\d{2}-\d{2}$/)) return date;
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}
