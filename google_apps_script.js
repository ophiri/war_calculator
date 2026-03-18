// ==== הדבק את הקוד הזה ב-Google Apps Script ====
// Extensions → Apps Script → מחק הכל → הדבק → Deploy

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date().toLocaleString("he-IL"),
    data.name,
    data.income,
    data.field,
    data.decrease + "%",
    data.area
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
