// Agrega las 4 columnas de condiciones comerciales a la hoja Cotizaciones (sin borrar datos).
function migrarColumnasCondiciones() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizaciones");
  if (!sheet) { SpreadsheetApp.getUi().alert("Hoja Cotizaciones no encontrada."); return; }

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const needed  = ["forma_pago", "plazo_entrega", "validez_oferta", "no_incluye"];
  let added = 0;
  needed.forEach(col => {
    if (!headers.includes(col)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col);
      added++;
    }
  });
  SpreadsheetApp.getUi().alert(
    added > 0
      ? "✅ " + added + " columna(s) agregada(s) a Cotizaciones."
      : "Las columnas ya existían. No se hicieron cambios."
  );
}

// Agrega la clave logo_id a Configuracion si no existe.
function migrarLogoConfig() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName("Configuracion");
  if (!cfg) { SpreadsheetApp.getUi().alert("Hoja Configuracion no encontrada."); return; }

  const claves = cfg.getDataRange().getValues().slice(1).map(r => String(r[0]).trim());
  if (!claves.includes("logo_id")) {
    cfg.appendRow(["logo_id", "", "ID del archivo de logo en Google Drive (aparece en el PDF de la cotización)"]);
    SpreadsheetApp.getUi().alert("✅ Clave logo_id agregada a Configuracion.\n\nSube tu logo a Drive, copia el ID del archivo y pégalo en la columna B de esa fila.");
  } else {
    SpreadsheetApp.getUi().alert("La clave logo_id ya existe.");
  }
}

// Agrega carpeta_firma y firma_id a Configuracion si no existen.
function migrarConfigFirma() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName("Configuracion");
  if (!cfg) { SpreadsheetApp.getUi().alert("Hoja Configuracion no encontrada."); return; }

  const claves = cfg.getDataRange().getValues().slice(1).map(r => String(r[0]).trim());
  let added = 0;
  if (!claves.includes("carpeta_firma")) {
    cfg.appendRow(["carpeta_firma", "", "ID de la carpeta de Drive donde se guarda la firma digital"]);
    added++;
  }
  if (!claves.includes("firma_id")) {
    cfg.appendRow(["firma_id", "", "ID del archivo de firma digital en Drive (se genera automáticamente)"]);
    added++;
  }
  SpreadsheetApp.getUi().alert(
    added > 0
      ? "✅ " + added + " clave(s) agregada(s) a Configuracion.\n\nCrea una carpeta en Drive para tu firma, copia su ID y pégalo en la columna B de 'carpeta_firma'."
      : "Las claves ya existían. No se hicieron cambios."
  );
}

function migrarDireccionAPU() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU");
  if (!sheet) { SpreadsheetApp.getUi().alert("Hoja APU no encontrada."); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  if (headers.includes("direccion")) {
    SpreadsheetApp.getUi().alert("La columna 'direccion' ya existe en APU."); return;
  }

  const clienteIdx = headers.indexOf("cliente");
  if (clienteIdx >= 0) {
    sheet.insertColumnAfter(clienteIdx + 1);
    sheet.getRange(1, clienteIdx + 2).setValue("direccion");
  } else {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("direccion");
  }
  SpreadsheetApp.getUi().alert("✅ Columna 'direccion' agregada a la hoja APU.\n\nVuelve a correr 'formatearHojas' para aplicar el formato.");
}

function migrarColumnaIVA() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizaciones");
  if (!sheet) { SpreadsheetApp.getUi().alert("Hoja Cotizaciones no encontrada."); return; }

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  if (headers.includes("iva_pct")) {
    SpreadsheetApp.getUi().alert("La columna iva_pct ya existe. No se requiere migración.");
    return;
  }

  const utilIdx = headers.indexOf("utilidad_pct");
  if (utilIdx < 0) { SpreadsheetApp.getUi().alert("No se encontró la columna utilidad_pct."); return; }

  sheet.insertColumnAfter(utilIdx + 1);
  const newCol = utilIdx + 2;
  sheet.getRange(1, newCol).setValue("iva_pct");

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, newCol, lastRow - 1, 1).setValue(19);

  SpreadsheetApp.getUi().alert("✅ Columna iva_pct agregada con 19% por defecto en todas las cotizaciones existentes.");
}

function limpiarDatos() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "⚠️ Limpiar datos",
    "Esto borrará TODOS los APUs y cotizaciones guardados.\n\n" +
    "La BD de precios y la configuración NO se tocarán.\n\n" +
    "¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const hojas  = ["APU", "APU_Items", "Cotizaciones", "Cotizacion_Items"];

  hojas.forEach(nombre => {
    const sheet = ss.getSheetByName(nombre);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  });

  ui.alert("✅ Datos limpiados. Las hojas APU, APU_Items, Cotizaciones y Cotizacion_Items están vacías.");
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // === BASE DE DATOS DE PRECIOS ===
  createSheet(ss, "Equipos", [
    "id", "nombre", "tarifa_dia", "partida"
  ]);

  createSheet(ss, "Materiales", [
    "id", "codigo", "categoria", "nombre", "unidad",
    "precio_sin_iva", "precio_con_iva", "precio_2026",
    "proveedor", "fecha_actualizacion", "partida"
  ]);

  createSheet(ss, "ManoObra", [
    "id", "descripcion", "salario_mensual",
    "prestaciones_pct", "costo_dia"
  ]);

  createSheet(ss, "Otros", [
    "id", "descripcion", "unidad", "costo_unitario"
  ]);

  // === APU ===
  createSheet(ss, "APU", [
    "id", "codigo_item", "descripcion", "unidad", "cliente", "direccion",
    "actividad", "subtotal_equipos", "subtotal_materiales",
    "subtotal_mano_obra", "subtotal_otros", "costo_neto",
    "administracion_pct", "imprevistos_pct", "utilidad_pct", "iva_pct", "valor_total",
    "fecha"
  ]);

  createSheet(ss, "APU_Items", [
    "id", "apu_id", "tipo", "recurso_id", "descripcion_manual",
    "cantidad", "rendimiento", "precio_unitario", "valor_parcial", "partida"
  ]);

  // === COTIZACIONES ===
  createSheet(ss, "Cotizaciones", [
    "id", "numero_oferta", "cliente", "direccion", "fecha",
    "valor_neto", "administracion_pct", "imprevistos_pct",
    "utilidad_pct", "iva_pct", "valor_total", "aprobada", "notas",
    "forma_pago", "plazo_entrega", "validez_oferta", "no_incluye"
  ]);

  createSheet(ss, "Cotizacion_Items", [
    "id", "cotizacion_id", "apu_id", "item_num", "descripcion",
    "unidad", "cantidad", "precio_apu", "valor_total"
  ]);

  // === CONFIGURACIÓN ===
  createSheet(ss, "Configuracion", ["clave", "valor", "descripcion"]);
  const cfg = ss.getSheetByName("Configuracion");
  cfg.appendRow(["carpeta_cotizaciones", "", "ID de la carpeta de Drive donde se guardan los PDFs de cotizaciones"]);
  cfg.appendRow(["carpeta_apus",         "", "ID de la carpeta de Drive donde se guardarán los PDFs de APUs"]);
  cfg.appendRow(["nombre_remitente",     "", "Tu nombre completo (aparece como remitente del correo)"]);
  cfg.appendRow(["empresa",              "", "Nombre de tu empresa o razón social"]);
  cfg.appendRow(["logo_id",             "", "ID del archivo de logo en Google Drive (aparece en el PDF)"]);
  cfg.appendRow(["carpeta_firma",        "", "ID de la carpeta de Drive donde se guarda la firma digital"]);
  cfg.appendRow(["firma_id",            "", "ID del archivo de firma digital en Drive (se genera automáticamente)"]);
  // Formato de la hoja config
  cfg.setColumnWidth(1, 180);
  cfg.setColumnWidth(2, 340);
  cfg.setColumnWidth(3, 420);
  cfg.getRange(2, 3, 2, 1).setFontColor("#9e9e9e").setFontStyle("italic");

  seedMateriales(ss);
  SpreadsheetApp.flush();
  seedManoObra(ss);
  seedEquipos(ss);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    "✅ Base de datos creada y datos cargados.\n\n" +
    "Ahora corre la función «formatearHojas» para aplicar el formato visual."
  );
}
