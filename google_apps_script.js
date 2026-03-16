// Google Apps Script — Web API สำหรับอัปเดต Google Sheets
// วิธีติดตั้ง:
// 1. เปิด Google Sheets → Extensions → Apps Script
// 2. วาง code นี้ทั้งหมด แทนที่ code เดิม
// 3. คลิก Deploy → New deployment
// 4. เลือก Type: Web app
// 5. Execute as: Me, Who has access: Anyone
// 6. คลิก Deploy → Copy URL
// 7. นำ URL ไปใส่ใน clip_cutter.py (ตัวแปร APPS_SCRIPT_URL)

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(data.sheet_name || "Hoshi");
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Sheet not found: " + (data.sheet_name || "Hoshi")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // รับ updates เป็น array ของ {cell: "A2", value: "success"}
    var updates = data.updates || [];
    
    for (var i = 0; i < updates.length; i++) {
      var cell = updates[i].cell;
      var value = updates[i].value;
      sheet.getRange(cell).setValue(value);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      updated: updates.length
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ทดสอบ: เรียกจาก Apps Script editor
function testUpdate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hoshi");
  sheet.getRange("A2").setValue("test");
  Logger.log("Done");
}
