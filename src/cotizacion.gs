// ─── CUADRO DE PARTIDAS ───────────────────────────────────────────────────────
// Agrupa todos los APU_Items de los APUs en esta cotización por partida.
// Devuelve array de { item_num, partida, valor } ordenado según PARTIDAS_FIJAS.

function getPartidasCotizacion(cotId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const citems = sheetToObjects(ss, "Cotizacion_Items").filter(r => String(r.cotizacion_id) === String(cotId));
  if (!citems.length) return [];
  const apuIds = citems.map(r => String(r.apu_id));

  const allItems = sheetToObjects(ss, "APU_Items")
    .filter(i => apuIds.includes(String(i.apu_id)));

  const agrupado = {};
  allItems.forEach(item => {
    const p = (item.partida && String(item.partida).trim())
      || getPartidaRecurso(ss, item.tipo, item.recurso_id);
    agrupado[p] = (agrupado[p] || 0) + (parseFloat(item.valor_parcial) || 0);
  });

  let num = 1;
  const resultado = PARTIDAS_FIJAS
    .filter(p => (agrupado[p] || 0) > 0)
    .map(p => ({ item_num: num++, partida: p, valor: Math.round(agrupado[p]) }));

  Object.keys(agrupado).forEach(p => {
    if (!PARTIDAS_FIJAS.includes(p) && agrupado[p] > 0)
      resultado.push({ item_num: num++, partida: p, valor: Math.round(agrupado[p]) });
  });

  return resultado;
}

// ─── DESCARGAR PDF (base64, descarga directa en el navegador) ────────────────

function descargarPDFBase64(cotId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cot = getCotizacionCompleta(cotId);
  if (!cot) return { ok: false, error: "Cotización no encontrada" };

  const tmpName = "_cot_dl_tmp_";
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

  const pdfBytes = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getContent();

  ss.deleteSheet(tmp);

  const slug = (cot.cliente || "cotizacion")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-zA-Z0-9 ]/g, "").trim().replace(/\s+/g, "_");
  const fileName = (cot.numero_oferta || "OF") + "_" + slug + ".pdf";

  return { ok: true, base64: Utilities.base64Encode(pdfBytes), fileName };
}

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
  const admPct    = parseFloat(cot.administracion_pct) || 0;
  const impPct    = parseFloat(cot.imprevistos_pct)    || 0;
  const utilPct   = parseFloat(cot.utilidad_pct)       || 0;
  const ivaPct    = parseFloat(cot.iva_pct)            || 0;

  const cfg       = getConfig();
  const empresa   = (cfg["empresa"]          || "").trim();
  const remitente = (cfg["nombre_remitente"] || "").trim();
  const logoId    = (cfg["logo_id"]          || "").trim();

  const formaPago     = String(cot.forma_pago     || "").trim();
  const plazoEntrega  = String(cot.plazo_entrega  || "").trim();
  const validezOferta = String(cot.validez_oferta || "").trim();
  const noIncluye     = String(cot.no_incluye     || "").trim();

  const MONEY    = '"$"#,##0';
  const NEGRO    = "#1a1a1a";
  const GRIS_OSC = "#424242";
  const GRIS_MED = "#9e9e9e";
  const GRIS_CLR = "#f5f5f5";
  const BORDE    = "#dddddd";
  const NC       = 6;

  sheet.setColumnWidth(1, 45);
  sheet.setColumnWidth(2, 285);
  sheet.setColumnWidth(3, 68);
  sheet.setColumnWidth(4, 72);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 130);

  let r = 1;

  // ─── LOGO + ENCABEZADO ────────────────────────────────────────────────────────
  sheet.setRowHeight(r, 85);
  sheet.getRange(r, 1, 1, 2).merge().setValue("").setBackground("#ffffff");
  sheet.getRange(r, 3, 1, 4).merge()
    .setValue(empresa || "").setFontSize(16).setFontWeight("bold").setFontColor("#000000")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#ffffff");
  if (logoId) {
    try {
      const img = sheet.insertImage(DriveApp.getFileById(logoId).getBlob(), 1, r);
      img.setWidth(215).setHeight(78);
    } catch(e) {}
  }
  r++;

  // Fila cliente + dirección (con wrap para que no se corte)
  const clienteTexto   = cot.cliente   ? "CLIENTE:    " + cot.cliente   : "";
  const direccionTexto = cot.direccion ? "DIRECCIÓN:  " + cot.direccion : "";
  sheet.setRowHeight(r, 34);
  sheet.getRange(r, 1, 1, 3).merge()
    .setValue(clienteTexto)
    .setFontSize(9).setFontColor("#222222").setFontWeight("bold")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setWrap(true).setBackground("#f5f5f5");
  sheet.getRange(r, 4, 1, 3).merge()
    .setValue(direccionTexto)
    .setFontSize(9).setFontColor("#333333")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setWrap(true).setBackground("#f5f5f5");
  r++;

  // Datos de oferta (N° oferta, fecha)
  sheet.setRowHeight(r, 20);
  const infoOferta = [
    cot.numero_oferta ? "N° Oferta: " + cot.numero_oferta : null,
    cot.fecha         ? "Fecha: "     + cot.fecha         : null,
  ].filter(Boolean).join("     ");
  sheet.getRange(r, 1, 1, NC).merge()
    .setValue(infoOferta).setFontSize(9).setFontColor("#555555")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground("#ffffff");
  r++;

  // Separador
  sheet.setRowHeight(r, 4);
  sheet.getRange(r, 1, 1, NC).merge().setValue("")
    .setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBackground("#ffffff");
  r++;

  // ─── ENCABEZADO DE TABLA ──────────────────────────────────────────────────────
  sheet.setRowHeight(r, 26);
  sheet.getRange(r, 1, 1, NC)
    .setValues([["ÍTEM", "DESCRIPCIÓN", "UND", "CANT.", "VR. UNITARIO", "VR. TOTAL"]])
    .setBackground(NEGRO).setFontColor("#ffffff").setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  r++;

  // ─── ÍTEMS DETALLADOS POR APU ─────────────────────────────────────────────────
  const SEC = [
    { key: "equipos",    label: "EQUIPOS"      },
    { key: "materiales", label: "MATERIALES"   },
    { key: "mano_obra",  label: "MANO DE OBRA" },
    { key: "otros",      label: "OTROS"         },
  ];

  if (items.length === 0) {
    sheet.setRowHeight(r, 22);
    sheet.getRange(r, 1, 1, NC).merge()
      .setValue("Sin ítems").setFontSize(9).setFontColor("#888888")
      .setHorizontalAlignment("center").setBackground("#ffffff");
    r++;
  }

  items.forEach((item, idx) => {
    const precio = parseFloat(item.precio_apu)  || 0;
    const cant   = parseFloat(item.cantidad)    || 1;
    const total  = cant * precio;

    // Fila principal del APU
    sheet.setRowHeight(r, 24);
    sheet.getRange(r, 1, 1, NC)
      .setBackground(GRIS_OSC)
      .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 1).setValue(idx + 1)
      .setFontSize(9).setFontWeight("bold").setFontColor("#ffffff")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(r, 2).setValue((item.codigo_item ? "[" + item.codigo_item + "]  " : "") + (item.descripcion || ""))
      .setFontSize(9).setFontWeight("bold").setFontColor("#ffffff")
      .setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true);
    sheet.getRange(r, 3).setValue(item.unidad || "")
      .setFontSize(9).setFontColor("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(r, 4).setValue(cant)
      .setFontSize(9).setFontColor("#ffffff").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange(r, 5).setValue(precio).setNumberFormat(MONEY)
      .setFontSize(9).setFontColor("#ffffff").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange(r, 6).setValue(total).setNumberFormat(MONEY)
      .setFontSize(9).setFontWeight("bold").setFontColor("#ffffff")
      .setHorizontalAlignment("right").setVerticalAlignment("middle");
    r++;

    // Sub-secciones del APU
    SEC.forEach(sec => {
      const subitems = item[sec.key] || [];
      if (!subitems.length) return;

      // Encabezado de sección
      sheet.setRowHeight(r, 17);
      sheet.getRange(r, 1, 1, NC).merge()
        .setValue("  " + sec.label)
        .setBackground(GRIS_CLR).setFontColor("#333333")
        .setFontWeight("bold").setFontSize(8).setFontStyle("italic")
        .setHorizontalAlignment("left").setVerticalAlignment("middle")
        .setBorder(true, true, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
      r++;

      // Ítems de la sección
      subitems.forEach((it, i) => {
        sheet.setRowHeight(r, 17);
        const bg = i % 2 === 0 ? "#ffffff" : "#fafafa";
        const rend = (it.rendimiento !== null && it.rendimiento !== undefined && it.rendimiento !== "") ? it.rendimiento : "";
        const cantTexto = (it.cantidad != null ? String(it.cantidad) : "") + (rend ? "  /  " + rend : "");
        sheet.getRange(r, 1).setValue("").setBackground(bg)
          .setBorder(true, true, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 2).setValue("      " + (it.descripcion_manual || "—"))
          .setFontSize(8).setFontColor("#222222").setHorizontalAlignment("left")
          .setVerticalAlignment("middle").setWrap(true).setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 3).setValue(it.unidad || "")
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("center")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 4).setValue(cantTexto)
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 5).setValue(parseFloat(it.precio_unitario) || 0).setNumberFormat(MONEY)
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 6).setValue(parseFloat(it.valor_parcial) || 0).setNumberFormat(MONEY)
          .setFontSize(8).setFontWeight("bold").setFontColor("#111111").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        r++;
      });
    });

    // Fila de subtotal del APU
    sheet.setRowHeight(r, 20);
    sheet.getRange(r, 1, 1, 5).merge()
      .setValue("SUBTOTAL ÍT. " + (idx + 1))
      .setFontSize(8).setFontWeight("bold").setFontColor("#333333")
      .setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBackground(GRIS_CLR)
      .setBorder(true, true, true, null, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 6).setValue(total).setNumberFormat(MONEY)
      .setFontSize(8).setFontWeight("bold").setFontColor("#111111")
      .setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBackground(GRIS_CLR)
      .setBorder(true, null, true, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  // ─── TOTAL COSTOS DIRECTOS ────────────────────────────────────────────────────
  sheet.setRowHeight(r, 24);
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("TOTAL COSTOS DIRECTOS")
    .setFontSize(10).setFontWeight("bold").setFontColor("#ffffff")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(GRIS_OSC)
    .setBorder(true, true, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(r, 6).setValue(valorNeto).setNumberFormat(MONEY)
    .setFontSize(10).setFontWeight("bold").setFontColor("#ffffff")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(GRIS_OSC)
    .setBorder(true, null, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  r++;

  // ─── AIU + VALOR TOTAL ────────────────────────────────────────────────────────
  const admVal  = Math.round(valorNeto * admPct  / 100);
  const impVal  = Math.round(valorNeto * impPct  / 100);
  const utilVal = Math.round(valorNeto * utilPct / 100);
  const sinIVA  = valorNeto + admVal + impVal + utilVal;
  const ivaVal  = Math.round(sinIVA * ivaPct / 100);
  const total   = sinIVA + ivaVal;

  const aiuFilas = [
    admPct  > 0 ? ["ADMINISTRACIÓN ("  + admPct  + "%)", admVal]  : null,
    impPct  > 0 ? ["IMPREVISTOS ("     + impPct  + "%)", impVal]  : null,
    utilPct > 0 ? ["UTILIDAD ("        + utilPct + "%)", utilVal] : null,
    ivaPct  > 0 ? ["IVA ("             + ivaPct  + "%)", ivaVal]  : null,
  ].filter(Boolean);

  aiuFilas.forEach(([label, val]) => {
    sheet.setRowHeight(r, 20);
    sheet.getRange(r, 1, 1, 5).merge()
      .setValue(label).setFontSize(9).setFontColor("#333333")
      .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground("#ffffff")
      .setBorder(true, true, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 6).setValue(val).setNumberFormat(MONEY)
      .setFontSize(9).setFontColor("#333333")
      .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground("#ffffff")
      .setBorder(true, null, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  // Valor total oferta
  sheet.setRowHeight(r, 28);
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("VALOR TOTAL OFERTA")
    .setFontSize(12).setFontWeight("bold").setFontColor("#ffffff")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(NEGRO)
    .setBorder(true, true, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(r, 6).setValue(total).setNumberFormat(MONEY)
    .setFontSize(12).setFontWeight("bold").setFontColor("#ffffff")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(NEGRO)
    .setBorder(true, null, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  r++;

  // ─── CONDICIONES COMERCIALES ──────────────────────────────────────────────────
  if (formaPago || plazoEntrega || validezOferta || noIncluye) {
    r += 2;
    sheet.setRowHeight(r, 22);
    sheet.getRange(r, 1, 1, NC).merge()
      .setValue("CONDICIONES COMERCIALES")
      .setFontSize(10).setFontWeight("bold").setFontColor("#ffffff")
      .setBackground(GRIS_OSC)
      .setHorizontalAlignment("left").setVerticalAlignment("middle");
    r++;

    [
      formaPago     ? ["FORMA DE PAGO",       formaPago]     : null,
      plazoEntrega  ? ["PLAZO DE ENTREGA",     plazoEntrega]  : null,
      validezOferta ? ["VALIDEZ DE LA OFERTA", validezOferta] : null,
    ].filter(Boolean).forEach(([label, val]) => {
      sheet.setRowHeight(r, 20);
      sheet.getRange(r, 1, 1, 2).merge().setValue(label)
        .setFontSize(9).setFontWeight("bold").setFontColor("#333333")
        .setBackground(GRIS_CLR).setVerticalAlignment("middle");
      sheet.getRange(r, 3, 1, 4).merge().setValue(val)
        .setFontSize(9).setFontColor("#333333").setBackground("#ffffff")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("middle");
      r++;
    });

    if (noIncluye) {
      r++;
      sheet.setRowHeight(r, 18);
      sheet.getRange(r, 1, 1, NC).merge()
        .setValue("NUESTRA OFERTA NO INCLUYE:")
        .setFontSize(9).setFontWeight("bold").setFontColor("#333333");
      r++;
      sheet.getRange(r, 1, 1, NC).merge()
        .setValue(noIncluye).setFontSize(9).setFontColor("#555555")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top");
      sheet.setRowHeight(r, 60);
      r++;
    }
  }

  // ─── FIRMA ────────────────────────────────────────────────────────────────────
  r += 2;
  const firmaId = (cfg["firma_id"] || "").trim();
  if (firmaId) {
    try {
      sheet.setRowHeight(r, 65);
      const firmaImg = sheet.insertImage(DriveApp.getFileById(firmaId).getBlob(), 1, r);
      firmaImg.setWidth(200).setHeight(58);
    } catch(e) {}
    r++;
  }
  sheet.setRowHeight(r, 4);
  sheet.getRange(r, 1, 1, 3).merge()
    .setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  r++;
  const firmaLinea = [remitente, empresa].filter(Boolean).join("\n");
  if (firmaLinea) {
    sheet.setRowHeight(r, remitente && empresa ? 36 : 20);
    sheet.getRange(r, 1, 1, NC).merge()
      .setValue(firmaLinea).setFontSize(9).setFontWeight("bold")
      .setHorizontalAlignment("left").setVerticalAlignment("top")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
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
    datos.iva_pct !== undefined ? datos.iva_pct : 19,
    0, "No",
    datos.notas          || "",
    datos.forma_pago     || "",
    datos.plazo_entrega  || "",
    datos.validez_oferta || "",
    datos.no_incluye     || "",
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

  // Si Google Sheets guardó los % como decimal (ej: 0.5 en lugar de 50),
  // los normalizamos a enteros (0 < v < 1 → v * 100).
  ["administracion_pct", "imprevistos_pct", "utilidad_pct", "iva_pct"].forEach(k => {
    const v = parseFloat(cot[k]) || 0;
    cot[k] = (v > 0 && v < 1) ? Math.round(v * 100) : v;
  });

  const itemsSheet = ss.getSheetByName("Cotizacion_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();

  if (itemsData.length < 2) { cot.items = []; return cot; }

  const ih = itemsData[0];
  cot.items = itemsData.slice(1)
    .filter(r => r[ih.indexOf("cotizacion_id")] == cotId)
    .map(r => { const o = {}; ih.forEach((h, i) => o[h] = r[i]); return o; });

  // Adjuntar sub-ítems del APU a cada ítem de la cotización
  const apuItemsSheet = ss.getSheetByName("APU_Items");
  const apuItemsData  = apuItemsSheet.getDataRange().getValues();
  if (apuItemsData.length > 1) {
    const aih        = apuItemsData[0];
    const allApuItems = apuItemsData.slice(1).map(r => {
      const o = {};
      aih.forEach((h, i) => o[h] = r[i]);
      return o;
    });
    cot.items.forEach(item => {
      const subs      = allApuItems.filter(ai => String(ai.apu_id) === String(item.apu_id));
      item.equipos    = subs.filter(i => i.tipo === "EQUIPO");
      item.materiales = subs.filter(i => i.tipo === "MATERIAL");
      item.mano_obra  = subs.filter(i => i.tipo === "MANO_OBRA");
      item.otros      = subs.filter(i => i.tipo === "OTRO");
    });
  } else {
    cot.items.forEach(item => {
      item.equipos = []; item.materiales = []; item.mano_obra = []; item.otros = [];
    });
  }

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
      const campos = ["cliente","direccion","administracion_pct","imprevistos_pct","utilidad_pct","iva_pct","aprobada","notas","forma_pago","plazo_entrega","validez_oferta","no_incluye"];
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

// ─── ELIMINAR COTIZACIÓN ─────────────────────────────────────────────────────

function eliminarCotizacion(cotId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const itemsSheet = ss.getSheetByName("Cotizacion_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();
  const ih         = itemsData[0];
  for (let i = itemsData.length - 1; i >= 1; i--) {
    if (itemsData[i][ih.indexOf("cotizacion_id")] == cotId) itemsSheet.deleteRow(i + 1);
  }

  const cotSheet = ss.getSheetByName("Cotizaciones");
  const cotData  = cotSheet.getDataRange().getValues();
  for (let i = 1; i < cotData.length; i++) {
    if (cotData[i][0] == cotId) { cotSheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

// ─── ENVIAR COTIZACIÓN POR CORREO ─────────────────────────────────────────────

function enviarCotizacionEmail(cotId, emailDestino) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const cot = getCotizacionCompleta(cotId);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Generar PDF
    const tmpName = "_cot_email_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    llenarHojaCotizacion(tmp, cot);
    SpreadsheetApp.flush();

    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=pdf&gid=" + tmp.getSheetId()
      + "&portrait=true&fitw=true&size=letter"
      + "&gridlines=false&printtitle=false&sheetnames=false"
      + "&top_margin=0.75&bottom_margin=0.75&left_margin=0.75&right_margin=0.75";

    const blob = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
    }).getBlob().setName("Cotizacion_" + (cot.numero_oferta || cotId) + ".pdf");

    ss.deleteSheet(tmp);

    const cfg     = getConfig();
    const nombre  = (cfg["nombre_remitente"] || "").trim() || Session.getActiveUser().getEmail();
    const empresa = (cfg["empresa"] || "").trim();

    const total    = Math.round(parseFloat(cot.valor_total) || 0);
    const totalFmt = "$" + total.toLocaleString("es-CO");
    const firma    = empresa ? nombre + "\n" + empresa : nombre;
    const asunto   = (cot.numero_oferta || "Cotización") + " - " + (cot.cliente || "");

    const textPlano = "Estimado/a cliente,\n\n"
      + "Le hago llegar la cotización " + (cot.numero_oferta || "") + " solicitada.\n"
      + "Valor total: " + totalFmt + "\n\n"
      + "El documento se encuentra adjunto a este correo.\n\n"
      + "Quedo atento a sus comentarios.\n\n" + firma;

    const firmaHtml = empresa
      ? '<strong>' + nombre + '</strong><br><span style="color:#666">' + empresa + '</span>'
      : '<strong>' + nombre + '</strong>';

    const htmlBody = '<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;color:#222">'
      + '<p>Estimado/a cliente,</p>'
      + '<p>Le hago llegar la cotización <strong>' + (cot.numero_oferta || "") + '</strong> solicitada.</p>'
      + '<table style="width:100%;border-collapse:collapse;margin:20px 0;font-size:14px">'
      + '<tr style="background:#f0f2f5"><td style="padding:10px 14px;border:1px solid #ddd;color:#555;width:130px">N° Oferta</td>'
      + '<td style="padding:10px 14px;border:1px solid #ddd;font-weight:600">' + (cot.numero_oferta || "—") + '</td></tr>'
      + '<tr><td style="padding:10px 14px;border:1px solid #ddd;color:#555">Cliente</td>'
      + '<td style="padding:10px 14px;border:1px solid #ddd">' + (cot.cliente || "—") + '</td></tr>'
      + '<tr style="background:#f0f2f5"><td style="padding:10px 14px;border:1px solid #ddd;color:#555">Valor Total</td>'
      + '<td style="padding:10px 14px;border:1px solid #ddd;font-weight:700;color:#1a237e">' + totalFmt + '</td></tr>'
      + '</table>'
      + '<p>El documento se encuentra adjunto a este correo.</p>'
      + '<p>Quedo atento a sus comentarios.</p>'
      + '<br><p style="margin:0">' + firmaHtml + '</p>'
      + '</div>';

    blob.setContentType("application/pdf");

    GmailApp.sendEmail(emailDestino, asunto, textPlano, {
      name:        nombre,
      replyTo:     Session.getActiveUser().getEmail(),
      attachments: [blob],
      htmlBody,
    });
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ─── ENVIAR EMAIL + GUARDAR PDF EN DRIVE (una sola operación) ────────────────

function enviarYGuardarPDF(cotId, emailDestino) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const cot = getCotizacionCompleta(cotId);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Generar PDF una sola vez
    const tmpName = "_cot_enviar_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    llenarHojaCotizacion(tmp, cot);
    SpreadsheetApp.flush();

    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=pdf&gid=" + tmp.getSheetId()
      + "&portrait=true&fitw=true&size=letter"
      + "&gridlines=false&printtitle=false&sheetnames=false"
      + "&top_margin=0.75&bottom_margin=0.75&left_margin=0.75&right_margin=0.75";

    const pdfBlob = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
    }).getBlob();

    ss.deleteSheet(tmp);

    // Nombre del archivo
    const now   = new Date();
    const meses = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];
    const dia   = String(now.getDate()).padStart(2, "0");
    const mes   = meses[now.getMonth()];
    const anio  = now.getFullYear();
    const hora  = String(now.getHours()).padStart(2, "0");
    const min   = String(now.getMinutes()).padStart(2, "0");
    const slug  = (cot.cliente || "sin_cliente")
      .normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/[^a-zA-Z0-9 ]/g, "").trim().replace(/\s+/g, "_");
    const fileName = `${dia}-${mes}-${anio}_${hora}-${min}_${slug}.pdf`;

    // Guardar en Drive
    const folder   = obtenerCarpetaPDF();
    const driveFile = folder.createFile(pdfBlob.setName(fileName));

    // Preparar y enviar correo con el mismo blob (clonar para no perder el nombre)
    const cfg     = getConfig();
    const nombre  = (cfg["nombre_remitente"] || "").trim() || Session.getActiveUser().getEmail();
    const empresa = (cfg["empresa"] || "").trim();

    const total    = Math.round(parseFloat(cot.valor_total) || 0);
    const totalFmt = "$" + total.toLocaleString("es-CO");
    const firma    = empresa ? nombre + "\n" + empresa : nombre;
    const asunto   = (cot.numero_oferta || "Cotización") + " - " + (cot.cliente || "");

    const textPlano = "Estimado/a cliente,\n\n"
      + "Le hago llegar la cotización " + (cot.numero_oferta || "") + " solicitada.\n"
      + "Valor total: " + totalFmt + "\n\n"
      + "El documento se encuentra adjunto a este correo.\n\n"
      + "Quedo atento a sus comentarios.\n\n" + firma;

    const firmaHtml = empresa
      ? '<strong>' + nombre + '</strong><br><span style="color:#666">' + empresa + '</span>'
      : '<strong>' + nombre + '</strong>';

    const htmlBody = '<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;color:#222">'
      + '<p>Estimado/a cliente,</p>'
      + '<p>Le hago llegar la cotización <strong>' + (cot.numero_oferta || "") + '</strong> solicitada.</p>'
      + '<table style="width:100%;border-collapse:collapse;margin:20px 0;font-size:14px">'
      + '<tr style="background:#f0f2f5"><td style="padding:10px 14px;border:1px solid #ddd;color:#555;width:130px">N° Oferta</td>'
      + '<td style="padding:10px 14px;border:1px solid #ddd;font-weight:600">' + (cot.numero_oferta || "—") + '</td></tr>'
      + '<tr><td style="padding:10px 14px;border:1px solid #ddd;color:#555">Cliente</td>'
      + '<td style="padding:10px 14px;border:1px solid #ddd">' + (cot.cliente || "—") + '</td></tr>'
      + '<tr style="background:#f0f2f5"><td style="padding:10px 14px;border:1px solid #ddd;color:#555">Valor Total</td>'
      + '<td style="padding:10px 14px;border:1px solid #ddd;font-weight:700;color:#1a237e">' + totalFmt + '</td></tr>'
      + '</table>'
      + '<p>El documento se encuentra adjunto a este correo.</p>'
      + '<p>Quedo atento a sus comentarios.</p>'
      + '<br><p style="margin:0">' + firmaHtml + '</p>'
      + '</div>';

    // Usar DriveFile como adjunto (evita clonar el blob)
    GmailApp.sendEmail(emailDestino, asunto, textPlano, {
      name:        nombre,
      replyTo:     Session.getActiveUser().getEmail(),
      attachments: [driveFile.getBlob().setName(fileName).setContentType("application/pdf")],
      htmlBody,
    });

    return { ok: true, nombre: fileName, url: driveFile.getUrl() };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ─── GUARDAR FIRMA DIGITAL ────────────────────────────────────────────────────

function guardarFirmaDigital(base64Png) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const base64 = base64Png.replace(/^data:image\/png;base64,/, "");
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), "image/png", "firma_digital.png");

  // Eliminar firma anterior si existe
  const cfg     = getConfig();
  const oldId   = (cfg["firma_id"] || "").trim();
  if (oldId) {
    try { DriveApp.getFileById(oldId).setTrashed(true); } catch(e) {}
  }

  // Guardar en carpeta_firma si está configurada, si no en la carpeta del spreadsheet
  const carpetaFirmaId = (cfg["carpeta_firma"] || "").trim();
  let folder;
  if (carpetaFirmaId) {
    try { folder = DriveApp.getFolderById(carpetaFirmaId); } catch(e) { folder = null; }
  }
  if (!folder) {
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  }
  const file = folder.createFile(blob);
  const fileId  = file.getId();

  // Actualizar firma_id en Configuracion
  const cfgSheet = ss.getSheetByName("Configuracion");
  const data     = cfgSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === "firma_id") {
      cfgSheet.getRange(i + 1, 2).setValue(fileId);
      return { ok: true };
    }
  }
  cfgSheet.appendRow(["firma_id", fileId, "ID del archivo de firma digital en Drive"]);
  return { ok: true };
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
      const iva    = parseFloat(cotData[i][cotH.indexOf("iva_pct")])            || 0;
      const sinIVA     = valorNeto * (1 + admin/100 + imprev/100 + util/100);
      const valorTotal = Math.round(sinIVA * (1 + iva/100));
      cotSheet.getRange(i + 1, cotH.indexOf("valor_neto")  + 1).setValue(Math.round(valorNeto));
      cotSheet.getRange(i + 1, cotH.indexOf("valor_total") + 1).setValue(valorTotal);
      break;
    }
  }
}
