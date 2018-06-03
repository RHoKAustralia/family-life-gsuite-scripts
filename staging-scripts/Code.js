var spreadsheetId = '1j9sZCLVUP1KWtWyctuLgrLpQ1SX_y1la2y34zS01JNg';

function logErrorAndThrow(errorMessage) {
  MailApp.sendEmail(email, 'Form submission failure notification', errorMessage);
  Logger.log(errorMessage);
  throw new Error(errorMessage);
}

function getColumnNumberByName(sheetId, columnName) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0]; // get the first sheet
  var maxColumn = sheet.getLastColumn();
  var firstRowData = sheet.getRange(1, 1, 1, maxColumn).getValues();
  
  for (var i = 0; i < firstRowData[0].length; i++) {
    if (firstRowData[0][i].toString().toLowerCase() == columnName.toString().toLowerCase()) return i;
  }
}


// identifies the logged-in user, matches the row of the master spreadsheet
// Email column must be 3rd column
// spreadsheet id must not change
function matchUserToRow(userEmail) {
  var sheet = SpreadsheetApp.openById(spreadsheetId); 
  var data = sheet.getDataRange().getValues();
  var emailIndex = getColumnNumberByName(spreadsheetId, 'Email address');
  for (var i = 0; i < data.length; i++) {
    if (data[i][emailIndex] == userEmail) { 
      Logger.log((i));
      return i + 1;
    }
  }
}

function getScaleItemByTitle(title) {
  var form = FormApp.getActiveForm();
  var items = form.getItems(FormApp.ItemType.SCALE);

  for (var i = 0; i < items.length; i++) {
    var scaleItem = items[i].asScaleItem();
    if (scaleItem.getTitle() == title) {
      return scaleItem
    }
  }
}

function getListItemByTitle(title) {
  var form = FormApp.getActiveForm();
  var items = form.getItems(FormApp.ItemType.LIST);
  Logger.log(items)

  for (var i = 0; i < items.length; i++) {
    Logger.log(items[i].getTitle())
    Logger.log(title)
    if (items[i].getTitle() == title) {
      Logger.log('found it!')
      return items[i]
    }
  }
}

function getOutReachCalendar() {
  // Attempt to get the Out Reach calendar if it exist 
  var outReachCalName = 'Family Life - Out Reach'
  var calendars = CalendarApp.getCalendarsByName(outReachCalName);
  if (calendars.length > 0) {
    return calendars[0]
  }

  return CalendarApp.createCalendar(outReachCalName)
}

function createOutReachEvent(duration_min, calendar) {
  var title = 'Out Reach event';
  var min_msec = 1000 * 60
  var start = new Date();
  var end = new Date(start.getTime() + min_msec * duration_min);
  var options = { location: 'outReachLoc', description: 'multipe visits', sendInvites: false };
  var event = calendar.createEvent(title, start, end, options)
    .setGuestsCanSeeGuests(false);
}

function createOutReachReminderEvent(duration_min, calendar, reminderOffset_min) {
  var title = 'Check back in';
  var eventDescription = 'Remember to use the status form to check back in'
  var now = new Date();
  var meetingDuration_msec = 1000 * 60 * 15
  var min_msec = 1000 * 60
  var start = new Date(now.getTime() + (min_msec * duration_min) + (reminderOffset_min * min_msec));
  var end = new Date(now.getTime() + (min_msec * duration_min) + (reminderOffset_min * min_msec) + meetingDuration_msec);
  var options = { description: eventDescription, sendInvites: false };
  var event = calendar.createEvent(title, start, end, options)
    .setGuestsCanSeeGuests(false)
    .addEmailReminder(5)
    .addPopupReminder(5)
}

function onFormSubmit(e) {
    // thank employee
    var form = FormApp.getActiveForm();
    var respondentEmail = e.response.getRespondentEmail();
    // form.setTitle(name + ", thanks for updating your status!");

    // get user-entered data from form submission
    var formItems = form.getItems();

    // first form item is "status"
    var status = e.response.getResponseForItem(formItems[0]).getResponse();

    // second form item is "notes"
    var notes = e.response.getResponseForItem(formItems[1]).getResponse();

    // get user's row, and cells to update
    var alphabet = [ 'A', 'B', 'C', 'D', 'E',
                    'F', 'G', 'H', 'I', 'J',
                    'K', 'L', 'M', 'N', 'O',
                    'P', 'Q', 'R', 'S', 'T',
                    'U', 'V', 'W', 'X', 'Y',
                    'Z', 'AA', 'AB', 'AC', 'AD' ];
                   
    var row = matchUserToRow(respondentEmail);
    if(!row) {
      logErrorAndThrow('Email address is not in spreadsheet');
    }
  
    var statusIndex = getColumnNumberByName(spreadsheetId, 'Status');
    var statusCell = alphabet[statusIndex] + row; 
    var notesIndex = getColumnNumberByName(spreadsheetId, 'Notes');
    var notesCell = alphabet[notesIndex] + row; 

    // update cells
    // assumes master spreadsheet id will not change! This is the id of 'Copy of In/out Board'
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    spreadsheet.getRange(statusCell).setValue(status);
    spreadsheet.getRange(notesCell).setValue(notes);

    updateCalendar(e.response);
}

function updateCalendar(formResponse) {
  var form = FormApp.getActiveForm();
  var itemResponses = formResponse.getItemResponses();

  var statusItem = getListItemByTitle('Current status');
  var statusResponse = formResponse.getResponseForItem(statusItem).getResponse();
  //The response string should be hard coded
  if (statusResponse == 'Home visit') {
    var minInHour = 60;

    var StatusResponseValue = formResponse.getItemResponses()[2].getResponse();
    var timeUntilBack_min_Value = StatusResponseValue * minInHour;

    var outReachCalendar = getOutReachCalendar();
    // create event
    var reminderOffset_min = 30;
    createOutReachEvent(timeUntilBack_min_Value, outReachCalendar);
    createOutReachReminderEvent(timeUntilBack_min_Value, outReachCalendar, reminderOffset_min);
  }
}
