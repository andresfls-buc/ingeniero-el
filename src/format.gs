// ─── FORMATEO VISUAL DE HOJAS ─────────────────────────────────────────────────
// Punto de entrada principal — también se puede correr manualmente desde el editor.

function formatearHojas() {
  formatearTodasLasHojas(SpreadsheetApp.getActiveSpreadsheet());
}

function formatearTodasLasHojas(ss) {
  const COP = '"$"#,##0';

  // anchos: ancho en píxeles de cada columna
  // precios: índices (1-based) de columnas numéricas/COP
  // textos:  índices de columnas de descripción (wrap + izquierda)
  fmtHoja(ss, "Equipos",
    [50, 320, 130],
    [3], [2]);

  fmtHoja(ss, "Materiales",
    [50, 85, 155, 290, 65, 130, 130, 110, 160, 115],
    [6, 7, 8], [4]);

  fmtHoja(ss, "ManoObra",
    [50, 210, 145, 100, 115],
    [3, 5], [2]);

  fmtHoja(ss, "Otros",
    [50, 260, 80, 130],
    [4], [2]);

  fmtHoja(ss, "APU",
    [50, 95, 230, 70, 190, 210, 115, 115, 115, 105, 125, 95],
    [7, 8, 9, 10, 11], [3, 5, 6]);

  fmtHoja(ss, "APU_Items",
    [50, 70, 100, 70, 260, 70, 90, 125, 125],
    [8, 9], [5]);

  fmtHoja(ss, "Cotizaciones",
    [50, 125, 210, 210, 100, 130, 85, 85, 85, 135, 75, 210],
    [6, 10], [3, 4, 12]);

  fmtHoja(ss, "Cotizacion_Items",
    [50, 110, 70, 70, 260, 70, 90, 135, 135],
    [8, 9], [5]);

  // Formato especial: columna "aprobada" en Cotizaciones
  aplicarCondicionalAprobada(ss);

  // Formato especial: columna "tipo" en APU_Items
  aplicarCondicionalTipo(ss);
}

// ─── FORMATEADOR GENÉRICO ─────────────────────────────────────────────────────

function fmtHoja(ss, nombre, anchos, colsPrecios, colsTexto) {
  const sheet = ss.getSheetByName(nombre);
  if (!sheet) return;

  const nCols  = anchos.length;
  const nData  = Math.max(sheet.getLastRow() - 1, 20);

  // ── 1. Encabezado ──────────────────────────────────────────────────────────
  const hdr = sheet.getRange(1, 1, 1, nCols);
  hdr.setBackground("#1a237e")
     .setFontColor("#ffffff")
     .setFontWeight("bold")
     .setFontSize(10)
     .setHorizontalAlignment("center")
     .setVerticalAlignment("middle")
     .setWrap(false);
  sheet.setRowHeight(1, 32);
  sheet.setFrozenRows(1);

  // ── 2. Filas alternas ──────────────────────────────────────────────────────
  const dataRange = sheet.getRange(2, 1, nData, nCols);

  // Colores alternos: par = blanco, impar = azul muy claro
  const bgs = Array.from({ length: nData }, (_, i) =>
    Array(nCols).fill(i % 2 === 0 ? "#ffffff" : "#e8eaf6")
  );
  dataRange.setBackgrounds(bgs);
  dataRange.setFontSize(10)
           .setFontColor("#212121")
           .setVerticalAlignment("middle");
  sheet.setRowHeights(2, nData, 22);

  // ── 3. Columna ID (col 1): gris y centrado ─────────────────────────────────
  sheet.getRange(2, 1, nData, 1)
       .setFontColor("#9e9e9e")
       .setHorizontalAlignment("center");

  // ── 4. Columnas de precio: formato COP y alineación derecha ───────────────
  (colsPrecios || []).forEach(col => {
    sheet.getRange(2, col, nData, 1)
         .setNumberFormat('"$"#,##0')
         .setHorizontalAlignment("right")
         .setFontWeight("bold")
         .setFontColor("#1b5e20");
  });

  // ── 5. Columnas de texto largo: wrap ──────────────────────────────────────
  (colsTexto || []).forEach(col => {
    sheet.getRange(2, col, nData, 1)
         .setWrap(true)
         .setHorizontalAlignment("left");
  });

  // ── 6. Ancho de columnas ──────────────────────────────────────────────────
  anchos.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // ── 7. Borde exterior e interno del encabezado ───────────────────────────
  hdr.setBorder(
    true, true, true, true, true, true,
    "#3949ab", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // ── 8. Borde ligero para datos ────────────────────────────────────────────
  if (nData > 0) {
    sheet.getRange(2, 1, nData, nCols).setBorder(
      true, true, true, true, true, true,
      "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID
    );
  }
}

// ─── FORMATO CONDICIONAL: APROBADA ────────────────────────────────────────────

function aplicarCondicionalAprobada(ss) {
  const sheet = ss.getSheetByName("Cotizaciones");
  if (!sheet) return;

  // Columna 11 = "aprobada"
  const rango = sheet.getRange(2, 11, Math.max(sheet.getLastRow(), 100), 1);

  // Limpiar reglas anteriores de esta hoja
  sheet.clearConditionalFormatRules();

  const reglaAprobada = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Sí")
    .setBackground("#e8f5e9")
    .setFontColor("#2e7d32")
    .setBold(true)
    .setRanges([rango])
    .build();

  const reglaPendiente = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("No")
    .setBackground("#fce4ec")
    .setFontColor("#c62828")
    .setRanges([rango])
    .build();

  sheet.setConditionalFormatRules([reglaAprobada, reglaPendiente]);
}

// ─── FORMATO CONDICIONAL: TIPO EN APU_ITEMS ───────────────────────────────────

function aplicarCondicionalTipo(ss) {
  const sheet = ss.getSheetByName("APU_Items");
  if (!sheet) return;

  const rango = sheet.getRange(2, 3, Math.max(sheet.getLastRow(), 100), 1);

  sheet.clearConditionalFormatRules();

  const colores = {
    EQUIPO:    { bg: "#e3f2fd", fg: "#0d47a1" },
    MATERIAL:  { bg: "#e8f5e9", fg: "#1b5e20" },
    MANO_OBRA: { bg: "#fff3e0", fg: "#e65100" },
    OTRO:      { bg: "#eceff1", fg: "#546e7a" },
  };

  const reglas = Object.entries(colores).map(([tipo, { bg, fg }]) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(tipo)
      .setBackground(bg)
      .setFontColor(fg)
      .setBold(true)
      .setRanges([rango])
      .build()
  );

  sheet.setConditionalFormatRules(reglas);
}
