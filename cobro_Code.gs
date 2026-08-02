// ═══════════════════════════════════════════════════════════════════════════
// SEVERA MOTOS · Google Apps Script — Backend del formulario de cobros
// ═══════════════════════════════════════════════════════════════════════════
//
// PASOS PARA PUBLICAR:
//  1. Ve a script.google.com → Nuevo proyecto → llámalo "SeveraMotos-Cobros"
//  2. Crea DOS archivos en el proyecto:
//       • Code.gs  ← este archivo
//       • index.html  ← el otro archivo que te di
//  3. En Code.gs, reemplaza SHEET_ID con el ID de tu Google Sheet
//     (el ID está en la URL: spreadsheets/d/ESTE_ES_EL_ID/edit)
//  4. Guarda todo (Ctrl+S)
//  5. Implementar → Nueva implementación → tipo: Aplicación web
//       Ejecutar como: Yo
//       Acceso: Cualquier persona
//  6. Copia la URL que aparece → es la que le mandas al cobrador
//
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_ID   = "PEGA_TU_GOOGLE_SHEET_ID_AQUI"; // ← CAMBIA ESTO
const SHEET_NAME = "Registro Diario";

// ─── Sirve el formulario HTML ───────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("SEVERA MOTOS · Registrar Cobro")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Graba el pago en Google Sheets ────────────────────────────────────────
function registrarPago(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let ws = ss.getSheetByName(SHEET_NAME);

  // Crea la hoja si no existe (primera vez)
  if (!ws) {
    ws = ss.insertSheet(SHEET_NAME);
    ws.appendRow([
      "fecha","cliente","moto","placa","","","",
      "pago_diario","pago_recibido","medio_pago","saldo",
      "","","","","observaciones"
    ]);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,16).setFontWeight("bold");
  }

  // 16 columnas A–P — mismo formato que el Excel local
  const row = new Array(16).fill("");
  row[0]  = new Date(data.fecha + "T12:00:00"); // A  fecha
  row[1]  = data.cliente;                        // B  cliente
  row[2]  = data.moto;                           // C  moto
  row[3]  = data.placa;                          // D  placa
  row[7]  = Number(data.tarifa);                 // H  pago_diario  (tarifa del cliente)
  row[8]  = Number(data.valor);                  // I  pago_recibido
  row[9]  = data.medio;                          // J  medio_pago
  row[15] = data.obs || "";                      // P  observaciones

  ws.appendRow(row);

  // Log para auditoría interna
  Logger.log("✅ Cobro: " + data.cliente + " | $" + data.valor + " | " + data.fecha);

  return { ok: true };
}
