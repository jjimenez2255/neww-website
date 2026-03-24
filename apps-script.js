function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Parse form data
    var name = e.parameter.name || '';
    var email = e.parameter.email || '';
    var project = e.parameter.project || '';
    var source = e.parameter.source || 'unknown';
    var timestamp = new Date().toLocaleString('en-US', {timeZone: 'America/Los_Angeles'});
    
    // Append row
    sheet.appendRow([timestamp, name, email, project, source]);
    
    return ContentService.createTextOutput(JSON.stringify({status: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('NEWW Agency Form Handler - POST only');
}
