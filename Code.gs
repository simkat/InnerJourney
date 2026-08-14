// Inner Journey Practice Log - Google Apps Script Backend
// Deploy as a Web App (Execute as: your account, Who has access: Anyone)

const SECRET = "pray";  // MUST match CONFIG.SECRET in index.html
const SHEET_NAME = "entries";  // Name of your sheet storing the data

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== SHEET_NAME) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] !== undefined ? row[i] : "");
      return obj;
    });
    
    return ContentService.createTextOutput(JSON.stringify({ ok: true, rows: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // Validate secret
    if (data.secret !== SECRET) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "Invalid secret" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const action = data.action;
    
    if (action === "add") {
      return addRow(sheet, data.row);
    } else if (action === "update") {
      return updateRow(sheet, data.id, data.fields);
    } else if (action === "delete") {
      return deleteRow(sheet, data.id);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "Unknown action" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function addRow(sheet, row) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = headers.map(h => row[h] !== undefined ? row[h] : "");
    sheet.appendRow(values);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, message: "Row added" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function updateRow(sheet, id, fields) {
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("id");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === id) {
        // Update this row
        Object.keys(fields).forEach(key => {
          const colIndex = headers.indexOf(key);
          if (colIndex !== -1) {
            sheet.getRange(i + 1, colIndex + 1).setValue(fields[key]);
          }
        });
        return ContentService.createTextOutput(JSON.stringify({ ok: true, message: "Row updated" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "ID not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function deleteRow(sheet, id) {
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("id");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === id) {
        sheet.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({ ok: true, message: "Row deleted" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "ID not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
