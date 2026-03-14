function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  // Add header row if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Guests', 'Events', 'City', 'Accommodation', 'Message']);
    // Bold the header
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  sheet.appendRow([
    new Date().toLocaleString('en-IN', {timeZone: 'Asia/Kolkata'}),
    data.name,
    data.phone,
    data.guests,
    data.events,
    data.city || '-',
    data.accommodation,
    data.message || '-'
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({status: 'success'}))
    .setMimeType(ContentService.MimeType.JSON);
}
