// Function triggered by form submission
function onFormSubmit(e) {

    var sheetName = "Form Responses 1";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var activeRow = e.range.getRow();
  
    // Fetch values directly from columns A to E in a single line
    var [columnA, columnB, columnC, columnD, columnE, columnF, columnG, columnH, columnI, columnJ, 
        columnK, columnL, columnM, columnN, columnO, columnP, columnQ, columnR, columnS, 
        columnT, columnU, columnV, columnW, columnX, columnY, columnZ
    ] = sheet.getRange(activeRow, 1, 1, 26).getValues()[0];
  
    var duration = calculateDateDifference(columnD, columnE);    // Calculate the duration between "From Date" (Column D) and "To Date" (Column E)
    sheet.getRange(activeRow, 6).setValue(duration);// Set the calculated duration in Column F (index 6)
  

    //Function for Calculate Date form From Date and To Date
  function calculateDateDifference(startDate, endDate) {
    var startTimestamp = new Date(startDate).getTime();
    var endTimestamp = new Date(endDate).getTime();
    if (isNaN(startTimestamp) || isNaN(endTimestamp) || startTimestamp > endTimestamp) {
      return "Invalid Date";
    }
    var millisecondsInADay = 1000 * 60 * 60 * 24;
    var differenceInDays = Math.floor((endTimestamp - startTimestamp) / millisecondsInADay)+1;
    return differenceInDays;
  } 


  
  var leaveBalance = getLeaveBalance(columnB, columnC); // Calculate leave balance based on the leave type (Column C)
  // Check if leave duration exceeds leave balance
  if (duration > leaveBalance) {
    // Send one type of email notification (insufficient balance)
    sendInsufficientBalanceEmail(columnB, columnC, duration);
    sheet.getRange(activeRow, 7).setValue("Reject-Auto");
  } else {
    // Send another type of email notification (sufficient balance)
    sendSufficientBalanceEmail(columnB, columnC, duration);
  }
    
}

function getLeaveBalance(email, leaveType) {
  var leaveBalanceSheetName = "LeaveBalance"; // Name of the "LeaveBalance" sheet
  var leaveBalanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    leaveBalanceSheetName
  );
  var leaveBalanceData = leaveBalanceSheet.getDataRange().getValues();

  // Find the row that matches the email and leave type
  for (var i = 1; i < leaveBalanceData.length; i++) {
    // Assuming row 1 contains headers
    var sheetEmail = leaveBalanceData[i][1]; // Email is in the second column
    var sheetLeaveType = leaveBalanceData[i][2]; // Leave type is in the third column
    var sheetLeaveBalance = leaveBalanceData[i][5]; // Leave balance is in the fifth column

    if (sheetEmail === email && sheetLeaveType === leaveType) {
      return sheetLeaveBalance;
      Logger.log(sheetLeaveBalance);
    }
  }

  // If no matching record is found, return a default value or handle as needed
  return 0; // Default balance if no match is found
}

// Function to send email notification for sufficient balance
function sendSufficientBalanceEmail(sendTo, leaveType, duration) {
  var mailSubject = "Leave Request Submitted - " + leaveType;
  var mailBody =
    "Dear Employee,<br>" +
    "Your request for " + leaveType + " leave has been submitted. The requested duration is " + duration +
    " days.<br>" +
    "Please ensure to manage your workload accordingly during your absence.<br><br>" +
    "Thank you,<br>" +
    "HR Department";

  sendMail(sendTo, mailSubject, mailBody);
}

// Function to send email notification for insufficient balance
function sendInsufficientBalanceEmail(sendTo, leaveType, duration) {
  var mailSubject = "Insufficient Leave Balance - " + leaveType;
  var mailBody =
    "Dear Employee,<br>" +
    "Your request for " + leaveType + " leave has been received, but the requested duration (" + duration +
    " days) exceeds your available leave balance.<br>" +
    "Please review your leave balance and consider adjusting your request accordingly.<br><br>" +
    "Thank you,<br>" +
    "HR Department";

  sendMail(sendTo, mailSubject, mailBody);
}

// Function triggered by column change (Approved/Reject)
function onColumnChangeApprovedReject(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var activeCell = e.source.getActiveRange();

  // Check if the changed cell is in column G (index 7)
  if (activeCell.getColumn() === 7) {
    var row = activeCell.getRow();

    var columnB = sheet.getRange(row, 2).getValue(); // Employee Email
    var columnC = sheet.getRange(row, 3).getValue(); // Leave Type
    var columnD = sheet.getRange(row, 4).getValue(); // Start Date
    var columnE = sheet.getRange(row, 5).getValue(); // End Date
    var columnChange = activeCell.getValue(); // Approved or Reject

    if (columnChange === "Approved") {
      var sendTo = columnB;
      var mailSubject = "Leave Application Approved - " + columnC;
      var mailBody =
        "Dear Employee,<br>" +
        "Your leave application for " + columnC + " from " + columnD + " to " + columnE + " has been approved.<br><br>" +
        "Thank you,<br>" +
        "HR Department";

      sendMail(sendTo, mailSubject, mailBody);

      // Create leave event in Google Calendar upon approval
      createEvents(columnB, new Date(columnD), new Date(columnE), columnC);
    } else if (columnChange === "Reject") {
      var sendTo = columnB;
      var mailSubject = "Leave Application Rejected - " + columnC;
      var mailBody =
        "Dear Employee,<br>" +
        "Your leave application for " + columnC + " from " + columnD + " to " + columnE + " has been rejected.<br><br>" +
        "Thank you,<br>" +
        "HR Department";

      sendMail(sendTo, mailSubject, mailBody);
    }
  }
}

// Function to send email
function sendMail(sendTo, mailSubject, mailBody) {
  // Implement your email sending code here (e.g., using MailApp)
  // This function should send the email to the specified recipient(s).
  MailApp.sendEmail({
    to: sendTo,
    subject: mailSubject,
    htmlBody: mailBody,
  });
}

// Function to create leave events in Google Calendar upon leave approval
function createEvents(employeeEmail, startDate, endDate, reason) {
  var calendarId = 'primary'; // Replace with your calendar ID or 'primary' for default calendar
  var event = {
    summary: 'Approved Leave: ' + reason,
    description: 'Approved leave for ' + employeeEmail + '. Reason: ' + reason,
    start: {
      dateTime: startDate.toISOString(),
      timeZone: 'Asia/Singapore' // Adjust based on your timezone
    },
    end: {
      dateTime: endDate.toISOString(),
      timeZone: 'Asia/Singapore' // Adjust based on your timezone
    },
    attendees: [
      { email: employeeEmail }
    ],
    reminders: {
      useDefault: false,
      overrides: [
        { method: 'email', minutes: 24 * 60 }, // Email reminder 24 hours before
        { method: 'popup', minutes: 10 } // Popup reminder 10 minutes before
      ]
    }
  };

  try {
    var createdEvent = Calendar.Events.insert(event, calendarId);
    Logger.log('Leave event created for ' + employeeEmail + ': ' + createdEvent.htmlLink);
  } catch (error) {
    Logger.log('Error creating leave event for ' + employeeEmail + ': ' + error.message);
  }
}

// Function to handle HTTP GET requests (for testing)
function doGet(e) {
  return HtmlService.createHtmlOutput('Hello, world!');
}

// Function to handle HTTP POST requests (for testing)
function doPost(e) {
  return ContentService.createTextOutput('Received POST request');
}



