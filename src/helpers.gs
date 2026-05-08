function createSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }

  sheet.appendRow(headers);
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
