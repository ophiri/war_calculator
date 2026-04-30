// ==== Google Apps Script — מחשבון פיצויים אמיתי ====
// ENDPOINT נפרד לחלוטין מהמחשבון-בדיחה!
// 
// הוראות התקנה:
// 1. צור Google Sheet חדש (נפרד מהשיט של הבדיחה)
// 2. Extensions → Apps Script
// 3. מחק הכל → הדבק את הקוד הזה
// 4. Deploy → New deployment → Web app
// 5. Execute as: Me, Who has access: Anyone
// 6. העתק את ה-URL ושים אותו ב-real_calculator.html בשורת REAL_CALC_SCRIPT_URL
//
// ===================================================

// חוסם בקשות GET
function doGet(e) {
  return ContentService.createTextOutput("Unauthorized");
}

// מקבל נתונים מהמחשבון האמיתי ושומר
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // בפעם הראשונה — יוצר כותרות
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "תאריך",
        "מסלול",
        "שם עסק",
        "ח.פ.",
        "אזור",
        "תחום",
        "עובדים",
        "פיצוי משוער (₪)",
        "זכאי",
        "מקור"
      ]);

      // עיצוב כותרות
      var headerRange = sheet.getRange(1, 1, 1, 10);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#0a3d62");
      headerRange.setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    // תרגום מסלול לעברית
    var trackMap = {
      "revenue": "מחזורים",
      "salary": "שכר עבודה",
      "direct": "נזק ישיר",
      "agriculture": "חקלאות"
    };

    sheet.appendRow([
      new Date().toLocaleString("he-IL"),
      trackMap[data.track] || data.track,
      data.bizName || "",
      data.bizId || "",
      data.area || "",
      data.sector || "",
      data.employees || 0,
      data.estimatedCompensation || 0,
      data.eligible ? "כן" : "לא",
      data.source || "real_calculator"
    ]);

    // עיצוב אוטומטי של עמודת הפיצוי כמספר
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 8).setNumberFormat("#,##0");

  } catch (err) {
    // שגיאה שקטה
    Logger.log("Error: " + err.message);
  }

  return ContentService.createTextOutput("");
}
