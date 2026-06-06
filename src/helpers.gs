function createSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    return;
  }

  // Only reset if there's no data (header-only or completely empty)
  if (sheet.getLastRow() < 2) {
    sheet.clear();
    sheet.appendRow(headers);
  }
}

// Lee la hoja "Configuracion" y devuelve un objeto { clave: valor }
function getConfig() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Configuracion");
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const cfg  = {};
  data.forEach(([clave, valor]) => {
    if (clave) cfg[String(clave).trim()] = String(valor || "").trim();
  });
  return cfg;
}

// Escribe un objeto { clave: valor } en la hoja Configuracion.
// Si la clave ya existe la actualiza; si no existe la agrega al final.
function guardarConfiguracion(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Configuracion");
  if (!sheet) return { ok: false, error: "Hoja Configuracion no encontrada" };

  const lastRow = sheet.getLastRow();
  const claves  = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim())
    : [];

  Object.entries(datos).forEach(([clave, valor]) => {
    const idx = claves.indexOf(clave);
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(valor);
    } else {
      sheet.appendRow([clave, valor, ""]);
      claves.push(clave);
    }
  });
  return { ok: true };
}
