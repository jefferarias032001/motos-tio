// ═══════════════════════════════════════════════════════════════════════════
// SEVERA MOTOS · Google Apps Script — Backend del formulario de cobros
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_ID   = "1b_d0lDAYCnPfbcCKRxzPuWN4NOS75Yw0VorFXUBOSPU";
const SHEET_NAME = "Registro Diario";

// Columnas del sheet (estructura compacta, sin columnas vacías):
// A(0):fecha  B(1):cliente  C(2):moto  D(3):placa
// E(4):pago_diario  F(5):pago_recibido  G(6):medio_pago  H(7):observaciones

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
      pago_diario:   Number(r[4]) || 0,
      pago_recibido: Number(r[5]) || 0,
      medio_pago:    String(r[6] || "").trim(),
      observaciones: String(r[7] || "").trim(),
    }));
}

// ─── Devuelve solo los registros de hoy ────────────────────────────────────
function obtenerRegistrosHoy() {
  const todos = obtenerRegistros();
  const hoy   = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd");
  return todos.filter(r => r.fecha === hoy);
}

// ─── Graba el pago en Google Sheets ────────────────────────────────────────
function registrarPago(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let ws = ss.getSheetByName(SHEET_NAME);

  if (!ws) {
    ws = ss.insertSheet(SHEET_NAME);
    ws.appendRow(["fecha","cliente","moto","placa","pago_diario","pago_recibido","medio_pago","observaciones"]);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,8).setFontWeight("bold");
  }

  ws.appendRow([
    data.fecha,
    data.cliente,
    data.moto,
    data.placa,
    Number(data.tarifa),
    Number(data.valor),
    data.medio,
    data.obs || ""
  ]);

  Logger.log("✅ Cobro: " + data.cliente + " | $" + data.valor + " | " + data.fecha);
  return { ok: true };
}
