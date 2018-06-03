var spreadsheetId = '1j9sZCLVUP1KWtWyctuLgrLpQ1SX_y1la2y34zS01JNg';

function getName() {
    var email = Session.getActiveUser().getEmail();
  
    var user = AdminDirectory.Users.get(email);
  
    if (!user) {
        Logger.log('Contact does not exist in the directory');
    }
  
    return user.name.givenName;
}

function preloadFormData() {
    var form = FormApp.getActiveForm();
    var name = getName();

    form.setTitle('Hello ' + name + ", please update your status");
}

function onOpen(e) {
    Logger.log('onOpenForm');
    preloadFormData();
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
function matchUserToRow() {
  var email = Session.getActiveUser().getEmail();
  var sheet = SpreadsheetApp.openById(spreadsheetId); 
  var data = sheet.getDataRange().getValues();
  var emailIndex = getColumnNumberByName(spreadsheetId, 'Email address');
  for (var i = 0; i < data.length; i++) {
    if (data[i][emailIndex] == email) { 
      Logger.log((i));
      return i + 1;
    }
  }
}

function onFormSubmit(e) {
    // thank employee
    var form = FormApp.getActiveForm();
    var name = getName();
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
                   
    var row = matchUserToRow();
    var statusIndex = getColumnNumberByName(spreadsheetId, 'Status');
    var statusCell = alphabet[statusIndex] + row; 
    var notesIndex = getColumnNumberByName(spreadsheetId, 'Notes');
    var notesCell = alphabet[notesIndex] + row; 

    // update cells
    // assumes master spreadsheet id will not change! This is the id of 'Copy of In/out Board'
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    spreadsheet.getRange(statusCell).setValue(status);
    spreadsheet.getRange(notesCell).setValue(notes);
}
