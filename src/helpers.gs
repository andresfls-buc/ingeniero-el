// Serializa una sección crítica que genera IDs (max+1) y escribe filas.
// SIN esto, dos llamadas concurrentes (ej: agregar varios APUs seguidos) leen el
// mismo max y asignan el MISMO id a dos filas → filas duplicadas/colisionadas que
// luego "desaparecen" al descargar y descuadran los totales. Con el lock, la
// segunda llamada espera a que la primera termine de escribir.
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // espera hasta 30s por el turno
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

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
    const k = String(clave || "").trim();
    if (!k) return;
    const v = String(valor || "").trim();
    // Si la clave está duplicada, NUNCA dejar que una fila vacía pise un valor real.
    if (cfg[k] === undefined || cfg[k] === "") cfg[k] = v;
  });
  return cfg;
}

// Limpieza de un solo uso: colapsa filas duplicadas de Configuracion a una por clave,
// conservando el valor no vacío y la descripción. Correr desde el editor una vez.
function deduplicarConfiguracion() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Configuracion");
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: "Hoja Configuracion vacía o inexistente." };

  const lastRow = sheet.getLastRow();
  const data    = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // clave, valor, descripcion
  const orden   = [];
  const map     = {};
  data.forEach(([clave, valor, desc]) => {
    const k = String(clave || "").trim();
    if (!k) return;
    const v = String(valor || "").trim();
    const d = String(desc  || "").trim();
    if (!map[k]) { map[k] = { valor: v, desc: d }; orden.push(k); }
    else {
      if (!map[k].valor && v) map[k].valor = v;   // preferir el valor no vacío
      if (!map[k].desc  && d) map[k].desc  = d;
    }
  });

  const filas = orden.map(k => [k, map[k].valor, map[k].desc]);
  sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  if (filas.length) sheet.getRange(2, 1, filas.length, 3).setValues(filas);
  return { ok: true, filasAntes: data.length, filasDespues: filas.length, claves: orden };
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
