// ==== הדבק את הקוד הזה ב-Google Apps Script ====
// Extensions → Apps Script → מחק הכל → הדבק → Deploy
//
// אבטחה:
// - אין doGet → אף אחד לא יכול לקרוא נתונים דרך הקישור
// - doPost רק שומר → הלקוח רק שולח, לא מקבל שום מידע בחזרה
// - הגיליון עצמו פרטי → רק אתה יכול לראות אותו ב-Google Sheets

// חוסם בקשות GET — אף אחד לא יכול לקרוא נתונים
function doGet(e) {
  return ContentService.createTextOutput("Unauthorized");
}

// מקבל נתונים ושומר — לא מחזיר שום מידע
function doPost(e) {
  try {
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
  } catch (err) {
    // שגיאה שקטה — לא חושפים שום מידע
  }

  // תמיד מחזיר תשובה ריקה — הלקוח לא יודע מה קרה
  return ContentService.createTextOutput("");
}
