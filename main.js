const SHEET_ID = "1Yz9XVDDtnoxW-ahMwflKDY4aQDl8eU7S8-uPdkhb9XI";
const SHEET_NAME = "DATA";

// Column map (0-based)
const COL = {
  TIMESTAMP: 0,
  NAME: 1,
  LASTNAME: 2,
  LASTMATER: 3,
  INSTITUTE: 4,
  POSITION: 5,
  EMAIL: 6,
  PHONE: 7,
  DAY1: 8,
  DAY2: 9,
  QR_ID: 10,
  CERT_SENT: 11,
};

// =======================
// LOAD HTML PAGE
// =======================
function doGet(e) {

  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : "";

  let file;

  if (page === "checkin") {
    file = HtmlService.createTemplateFromFile("checkin");
    file.day = e.parameter.day || "";
  }
  else if (page === "admin") {
    file = HtmlService.createTemplateFromFile("admin");
  }

  else if (page === "assistant") {
    file = HtmlService.createTemplateFromFile("assistant"); 
  }
  
  else if (page === "assistant2") {
    file = HtmlService.createTemplateFromFile("assistant-two");
  }

  else {
    file = HtmlService.createTemplateFromFile("register");
  }

 
  return file.evaluate()
    .setTitle("Sistema de Eventos")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// =======================
// REGISTER USER
// =======================
function saveUser(data) {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Hoja no encontrada. Verifica SHEET_NAME = DATA");
  }

  const rows = sheet.getDataRange().getValues();
  const inputEmail = data.email.toLowerCase().trim();

  // Prevent duplicate email
  for (let i = 1; i < rows.length; i++) {

    const sheetEmail = String(rows[i][COL.EMAIL]).toLowerCase().trim();

    if (sheetEmail === inputEmail) {
      return "El correo ya está registrado";
    }
  }

  const row = [];

  row[COL.TIMESTAMP] = new Date();
  row[COL.NAME] = data.name;
  row[COL.LASTNAME] = data.lastname;
  row[COL.LASTMATER] = data.lastmater;
  row[COL.INSTITUTE] = data.institute;
  row[COL.POSITION] = data.position;
  row[COL.EMAIL] = inputEmail;
  row[COL.PHONE] = data.phone;
  row[COL.DAY1] = "";
  row[COL.DAY2] = "";
  row[COL.QR_ID] = generateQRId();
  row[COL.CERT_SENT] = "";

  sheet.appendRow(row);

  return "Registrado correctamente";
}


// =======================
// CHECK-IN
// =======================
function markAttendance(email, day) {

  if (day !== "1" && day !== "2") {
    return "Día inválido";
  }

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const inputEmail = email.toLowerCase().trim();

  for (let i = 1; i < data.length; i++) {

    const sheetEmail = String(data[i][COL.EMAIL]).toLowerCase().trim();

    if (sheetEmail === inputEmail) {

      const col = day == "1" ? COL.DAY1 : COL.DAY2;

      // Already checked
      if (data[i][col] == 1) {
        return "Ya registraste asistencia para el día " + day;
      }

      // Mark attendance
      sheet.getRange(i + 1, col + 1).setValue(1);

      return "Asistencia registrada correctamente ✔ Día " + day;
    }
  }

  return "Usuario no encontrado";
}


// =======================
// GET DATA (ADMIN)
// =======================
function getData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  return sheet.getDataRange().getValues();
}


function saveAssistant(data) {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("ATTENDANCE");

  if (!sheet) {
    throw new Error("Hoja ATTENDANCE no encontrada");
  }

  const email = data.email.toLowerCase().trim();

  // Get existing emails (column C)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues();

    for (let i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).toLowerCase().trim() === email) {
        return "Este correo ya fue registrado";
      }
    }
  }

  // Save new row
  sheet.appendRow([
    new Date(),
    data.name,
    email
  ]);

  return "Asistente registrado correctamente";
}

function saveAssistantDay2(data) {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName("ATTENDANCE2");

  if (!sheet) {
    throw new Error("Hoja ATTENDANCE2 no encontrada");
  }

  const email = data.email.toLowerCase().trim();

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues();

    for (let i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).toLowerCase().trim() === email) {
        return "Este correo ya fue registrado (Día 2)";
      }
    }
  }

  sheet.appendRow([
    new Date(),
    data.name,
    email
  ]);

  return "Asistencia registrada correctamente ✔ Día 2";
}


// =======================
// QR ID GENERATOR
// =======================
function generateQRId() {
  return "ID-" + Math.random().toString(36).substring(2, 10);
}
