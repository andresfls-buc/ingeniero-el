// ─── EXPORTAR PDF ────────────────────────────────────────────────────────────

function exportarCotizacionPDF(cotId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cot = getCotizacionCompleta(cotId);
  if (!cot) return { ok: false, error: "Cotización no encontrada" };

  const tmpName = "_cot_pdf_";
  let tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  llenarHojaCotizacion(tmp, cot);
  SpreadsheetApp.flush();

  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId()
    + "/export?format=pdf&gid=" + tmp.getSheetId()
    + "&portrait=true&fitw=true&size=letter"
    + "&gridlines=false&printtitle=false&sheetnames=false"
    + "&top_margin=0.75&bottom_margin=0.75&left_margin=0.75&right_margin=0.75";

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  });

  const now    = new Date();
  const meses  = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];
  const dia    = String(now.getDate()).padStart(2, "0");
  const mes    = meses[now.getMonth()];
  const anio   = now.getFullYear();
  const hora   = String(now.getHours()).padStart(2, "0");
  const min    = String(now.getMinutes()).padStart(2, "0");
  const slug   = (cot.cliente || "sin_cliente")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-zA-Z0-9 ]/g, "").trim().replace(/\s+/g, "_");
  const fileName = `${dia}-${mes}-${anio}_${hora}-${min}_${slug}.pdf`;

  const folder = obtenerCarpetaPDF();
  const file   = folder.createFile(response.getBlob().setName(fileName));

  ss.deleteSheet(tmp);

  return { ok: true, url: file.getUrl(), nombre: fileName };
}

function obtenerCarpetaPDF() {
  const FOLDER_ID = (getConfig()["carpeta_cotizaciones"] || "").trim();
  if (!FOLDER_ID) throw new Error(
    "Configura el ID de tu carpeta de Drive en la hoja 'Configuracion' → fila 'carpeta_cotizaciones'. " +
    "Abre la carpeta en Drive, copia el ID del final de la URL y pégalo ahí."
  );
  const raiz = DriveApp.getFolderById(FOLDER_ID);

  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                  "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  const mes   = meses[new Date().getMonth()];

  const iter = raiz.getFoldersByName(mes);
  if (iter.hasNext()) return iter.next();
  return raiz.createFolder(mes);
}

function llenarHojaCotizacion(sheet, cot) {
  const items     = cot.items || [];
  const valorNeto = parseFloat(cot.valor_neto)  || 0;
  const total     = parseFloat(cot.valor_total) || 0;
  const admin     = parseFloat(cot.administracion_pct) || 0;
  const imprev    = parseFloat(cot.imprevistos_pct)    || 0;
  const util      = parseFloat(cot.utilidad_pct)       || 0;
  const MONEY     = '"$"#,##0';
  const AZUL      = "#1a237e";
  const AZUL_LITE = "#e8eaf6";

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 65);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);

  let r = 1;

  // Encabezado
  const h1 = sheet.getRange(r, 1, 1, 6);
  h1.merge().setValue("COTIZACIÓN / OFERTA DE SERVICIOS")
    .setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center")
    .setBackground(AZUL).setFontColor("#ffffff");
  sheet.setRowHeight(r, 36); r++;

  sheet.getRange(r, 1, 1, 3).merge().setValue("N° Oferta:  " + (cot.numero_oferta || "")).setFontWeight("bold");
  sheet.getRange(r, 4, 1, 3).merge().setValue("Fecha:  " + (cot.fecha || "")).setHorizontalAlignment("right");
  r++;

  sheet.getRange(r, 1, 1, 6).merge().setValue("Cliente:  " + (cot.cliente || "")).setFontWeight("bold");
  r++;

  if (cot.direccion) {
    sheet.getRange(r, 1, 1, 6).merge().setValue("Dirección:  " + cot.direccion);
    r++;
  }
  if (cot.notas) {
    sheet.getRange(r, 1, 1, 6).merge().setValue("Notas:  " + cot.notas)
      .setFontStyle("italic").setFontColor("#666666");
    r++;
  }
  r++; // espacio

  // Cabecera tabla
  sheet.getRange(r, 1, 1, 6).setValues([["#", "DESCRIPCIÓN", "UNIDAD", "CANT.", "PRECIO APU", "VALOR TOTAL"]])
    .setFontWeight("bold").setBackground(AZUL_LITE).setFontColor(AZUL).setHorizontalAlignment("center");
  r++;

  const firstItemRow = r;
  if (items.length === 0) {
    sheet.getRange(r, 1, 1, 6).merge().setValue("Sin ítems")
      .setFontColor("#888888").setHorizontalAlignment("center");
    r++;
  } else {
    items.forEach((item, idx) => {
      const cant   = parseFloat(item.cantidad)    || 1;
      const precio = parseFloat(item.precio_apu)  || 0;
      const vt     = parseFloat(item.valor_total) || cant * precio;
      if (idx % 2 === 0) sheet.getRange(r, 1, 1, 6).setBackground("#f8f9fa");
      sheet.getRange(r, 1).setValue(item.item_num || idx + 1).setHorizontalAlignment("center");
      sheet.getRange(r, 2).setValue(item.descripcion || "");
      sheet.getRange(r, 3).setValue(item.unidad || "").setHorizontalAlignment("center");
      sheet.getRange(r, 4).setValue(cant).setHorizontalAlignment("right");
      sheet.getRange(r, 5).setValue(precio).setNumberFormat(MONEY).setHorizontalAlignment("right");
      sheet.getRange(r, 6).setValue(vt).setNumberFormat(MONEY).setFontWeight("bold").setHorizontalAlignment("right");
      r++;
    });
  }

  // Costo neto
  sheet.getRange(r, 1, 1, 5).merge().setValue("COSTO NETO")
    .setFontWeight("bold").setHorizontalAlignment("right").setBackground(AZUL_LITE);
  sheet.getRange(r, 6).setValue(valorNeto).setNumberFormat(MONEY)
    .setFontWeight("bold").setBackground(AZUL_LITE).setHorizontalAlignment("right");
  r++; r++; // espacio

  // AIU
  [["Administración", admin], ["Imprevistos", imprev], ["Utilidad", util]].forEach(([label, pct]) => {
    sheet.getRange(r, 1, 1, 4).merge().setValue(label + "  (" + pct + "%)").setHorizontalAlignment("right");
    sheet.getRange(r, 5, 1, 1).merge();
    sheet.getRange(r, 6).setValue(valorNeto * pct / 100).setNumberFormat(MONEY).setHorizontalAlignment("right");
    r++;
  });
  r++;

  // Total final
  sheet.getRange(r, 1, 1, 5).merge().setValue("TOTAL OFERTA")
    .setFontSize(12).setFontWeight("bold").setHorizontalAlignment("right")
    .setBackground(AZUL).setFontColor("#ffffff");
  sheet.getRange(r, 6).setValue(total).setNumberFormat(MONEY)
    .setFontSize(12).setFontWeight("bold").setBackground(AZUL).setFontColor("#ffffff").setHorizontalAlignment("right");
  r++;

  if (cot.aprobada === "Sí") {
    r++;
    sheet.getRange(r, 1, 1, 6).merge().setValue("✓  COTIZACIÓN APROBADA")
      .setFontWeight("bold").setFontColor("#2e7d32").setHorizontalAlignment("center").setBackground("#e8f5e9");
  }

  // Borde en la tabla de ítems
  const numFilas = items.length + 2;
  sheet.getRange(firstItemRow - 1, 1, numFilas, 6)
    .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
}

// ─── LISTAR COTIZACIONES ──────────────────────────────────────────────────────

function listarCotizaciones() {
  return sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "Cotizaciones");
}

// ─── CREAR COTIZACIÓN ─────────────────────────────────────────────────────────

function crearCotizacion(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizaciones");
  const data  = sheet.getDataRange().getValues();

  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId  = lastId + 1;

  const year     = new Date().getFullYear();
  const existing = data.length > 1
    ? data.slice(1).filter(r => String(r[1] || "").startsWith("OF-" + year)).length
    : 0;
  const numero = `OF-${year}-${String(existing + 1).padStart(3, "0")}`;
  const fecha  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  sheet.appendRow([
    newId, numero,
    datos.cliente          || "",
    datos.direccion        || "",
    fecha,
    0,
    datos.administracion_pct || 0,
    datos.imprevistos_pct    || 0,
    datos.utilidad_pct       || 0,
    0, "No",
    datos.notas || ""
  ]);

  return { id: newId, numero_oferta: numero };
}

// ─── GET COTIZACIÓN COMPLETA ──────────────────────────────────────────────────

function getCotizacionCompleta(cotId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cotSheet = ss.getSheetByName("Cotizaciones");
  const cotData  = cotSheet.getDataRange().getValues();
  const cotH     = cotData[0];
  const cotRow   = cotData.slice(1).find(r => r[0] == cotId);
  if (!cotRow) return null;

  const cot = {};
  cotH.forEach((h, i) => cot[h] = cotRow[i]);

  const itemsSheet = ss.getSheetByName("Cotizacion_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();

  if (itemsData.length < 2) { cot.items = []; return cot; }

  const ih = itemsData[0];
  cot.items = itemsData.slice(1)
    .filter(r => r[ih.indexOf("cotizacion_id")] == cotId)
    .map(r => { const o = {}; ih.forEach((h, i) => o[h] = r[i]); return o; });

  return cot;
}

// ─── AGREGAR APU A COTIZACIÓN ─────────────────────────────────────────────────

function agregarAPUaCotizacion(cotId, apuId, cantidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const apuSheet = ss.getSheetByName("APU");
  const apuData  = apuSheet.getDataRange().getValues();
  const apuH     = apuData[0];
  const apuRow   = apuData.slice(1).find(r => r[0] == apuId);
  if (!apuRow) return { ok: false };

  const apu = {};
  apuH.forEach((h, i) => apu[h] = apuRow[i]);

  const sheet  = ss.getSheetByName("Cotizacion_Items");
  const data   = sheet.getDataRange().getValues();
  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId   = lastId + 1;
  const itemNum = data.length > 1
    ? data.slice(1).filter(r => r[1] == cotId).length + 1
    : 1;

  const cant      = parseFloat(cantidad) || 1;
  const precioAPU = parseFloat(apu.costo_neto) || 0;
  const valTotal  = cant * precioAPU;

  sheet.appendRow([
    newId, cotId, apuId, itemNum,
    apu.descripcion || apu.codigo_item || "",
    apu.unidad || "",
    cant, precioAPU, valTotal
  ]);

  recalcularCotizacion(ss, cotId);
  return { id: newId, valor_total: valTotal };
}

// ─── ELIMINAR ÍTEM DE COTIZACIÓN ─────────────────────────────────────────────

function eliminarItemCotizacion(itemId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const cotId = data[i][h.indexOf("cotizacion_id")];
      sheet.deleteRow(i + 1);
      recalcularCotizacion(ss, cotId);
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─── ACTUALIZAR CANTIDAD DE ÍTEM ─────────────────────────────────────────────

function actualizarCantidadItem(itemId, cantidad) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const row       = [...data[i]];
      const cotId     = row[h.indexOf("cotizacion_id")];
      const precioAPU = parseFloat(row[h.indexOf("precio_apu")]) || 0;
      const cant      = parseFloat(cantidad) || 1;
      row[h.indexOf("cantidad")]    = cant;
      row[h.indexOf("valor_total")] = cant * precioAPU;
      sheet.getRange(i + 1, 1, 1, h.length).setValues([row]);
      recalcularCotizacion(ss, cotId);
      return { ok: true, valor_total: cant * precioAPU };
    }
  }
  return { ok: false };
}

// ─── ACTUALIZAR DATOS DE COTIZACIÓN ──────────────────────────────────────────

function actualizarCotizacion(cotId, datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizaciones");
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == cotId) {
      const campos = ["cliente","direccion","administracion_pct","imprevistos_pct","utilidad_pct","aprobada","notas"];
      campos.forEach(campo => {
        if (datos[campo] !== undefined) {
          const col = h.indexOf(campo);
          if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(datos[campo]);
        }
      });
      recalcularCotizacion(ss, cotId);
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─── HELPER: RECALCULAR TOTALES ───────────────────────────────────────────────

function recalcularCotizacion(ss, cotId) {
  const itemsSheet = ss.getSheetByName("Cotizacion_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();

  let valorNeto = 0;
  if (itemsData.length > 1) {
    const ih = itemsData[0];
    valorNeto = itemsData.slice(1)
      .filter(r => r[ih.indexOf("cotizacion_id")] == cotId)
      .reduce((s, r) => s + (parseFloat(r[ih.indexOf("valor_total")]) || 0), 0);
  }

  const cotSheet = ss.getSheetByName("Cotizaciones");
  const cotData  = cotSheet.getDataRange().getValues();
  const cotH     = cotData[0];

  for (let i = 1; i < cotData.length; i++) {
    if (cotData[i][0] == cotId) {
      const admin  = parseFloat(cotData[i][cotH.indexOf("administracion_pct")]) || 0;
      const imprev = parseFloat(cotData[i][cotH.indexOf("imprevistos_pct")])    || 0;
      const util   = parseFloat(cotData[i][cotH.indexOf("utilidad_pct")])       || 0;
      const total  = valorNeto * (1 + (admin + imprev + util) / 100);
      cotSheet.getRange(i + 1, cotH.indexOf("valor_neto")  + 1).setValue(valorNeto);
      cotSheet.getRange(i + 1, cotH.indexOf("valor_total") + 1).setValue(total);
      break;
    }
  }
}
