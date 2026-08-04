// ═══════════════════════════════════════════════════════════════════════════
// SEVERA MOTOS · Google Apps Script — Backend del formulario de cobros
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_ID   = "1b_d0lDAYCnPfbcCKRxzPuWN4NOS75Yw0VorFXUBOSPU";
const SHEET_NAME = "Registro Diario";

// ─── Sirve el formulario O los datos según el parámetro ?action ─────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // ?action=data → devuelve los registros como JSON (para el dashboard)
  if (action === "data") {
    const filas = obtenerRegistros();
    return ContentService
      .createTextOutput(JSON.stringify(filas))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Sin parámetro → formulario para el cobrador
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("SEVERA MOTOS · Registrar Cobro")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Devuelve todos los registros del sheet como array de objetos ───────────
function obtenerRegistros() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName(SHEET_NAME);
  if (!ws) return [];

  const datos = ws.getDataRange().getValues();
  if (datos.length <= 1) return [];

  return datos.slice(1)
    .filter(r => r[0] && r[1])   // debe tener fecha y cliente
    .map(r => ({
      fecha:         Utilities.formatDate(new Date(r[0]), "America/Bogota", "yyyy-MM-dd"),
      cliente:       String(r[1] || "").trim(),
      moto:          String(r[2] || "").trim(),
      placa:         String(r[3] || "").trim(),
      pago_diario:   Number(r[7]) || 0,
      pago_recibido: Number(r[8]) || 0,
      medio_pago:    String(r[9] || "").trim(),
    }));
}

// ─── Graba el pago en Google Sheets ────────────────────────────────────────
function registrarPago(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let ws = ss.getSheetByName(SHEET_NAME);

  if (!ws) {
    ws = ss.insertSheet(SHEET_NAME);
    ws.appendRow(["fecha","cliente","moto","placa","","","",
      "pago_diario","pago_recibido","medio_pago","saldo","","","","","observaciones"]);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,16).setFontWeight("bold");
  }

  const row = new Array(16).fill("");
  row[0]  = data.fecha;
  row[1]  = data.cliente;
  row[2]  = data.moto;
  row[3]  = data.placa;
  row[7]  = Number(data.tarifa);
  row[8]  = Number(data.valor);
  row[9]  = data.medio;
  row[15] = data.obs || "";

  ws.appendRow(row);
  Logger.log("✅ Cobro: " + data.cliente + " | $" + data.valor + " | " + data.fecha);
  return { ok: true };
}
