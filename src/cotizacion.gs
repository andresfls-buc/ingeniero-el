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

function descargarPDFBase64(cotId, tipo) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const { cot, render } = _prepararCotizacion(cotId, tipo);
  if (!cot) return { ok: false, error: "Cotización no encontrada" };

  const tmpName = "_cot_dl_tmp_";
  let tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  render(tmp, cot);
  SpreadsheetApp.flush();

  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId()
    + "/export?format=xlsx&gid=" + tmp.getSheetId();

  const xlsxBytes = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getContent();

  ss.deleteSheet(tmp);

  const slug = (cot.cliente || "cotizacion")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-zA-Z0-9 ]/g, "").trim().replace(/\s+/g, "_");
  const fileName = (cot.numero_oferta || "OF") + "_" + slug + ".xlsx";

  return { ok: true, base64: Utilities.base64Encode(xlsxBytes), fileName };
}

// ─── EXPORTAR PDF ────────────────────────────────────────────────────────────

function exportarCotizacionPDF(cotId, tipo) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const { cot, render } = _prepararCotizacion(cotId, tipo);
  if (!cot) return { ok: false, error: "Cotización no encontrada" };

  const tmpName = "_cot_pdf_";
  let tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  render(tmp, cot);
  SpreadsheetApp.flush();

  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId()
    + "/export?format=xlsx&gid=" + tmp.getSheetId();

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
  const fileName = `${dia}-${mes}-${anio}_${hora}-${min}_${slug}.xlsx`;

  const folder = obtenerCarpetaPDF();
  const file   = folder.createFile(response.getBlob().setName(fileName));

  ss.deleteSheet(tmp);

  return { ok: true, url: file.getUrl(), nombre: fileName };
}

function obtenerCarpetaPDF() {
  const FOLDER_ID = (getConfig()["carpeta_cotizaciones_internas"] || "").trim();
  if (!FOLDER_ID) throw new Error(
    "Configura el ID de tu carpeta de Drive en la hoja 'Configuracion' → fila 'carpeta_cotizaciones_internas'. " +
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
  const NEGRO    = "#a8a8a8";
  const GRIS_OSC = "#c8c8c8";
  const GRIS_MED = "#9e9e9e";
  const GRIS_CLR = "#f5f5f5";
  const BORDE    = "#dddddd";
  const NC       = 6;

  const ssCot = SpreadsheetApp.getActiveSpreadsheet();

  // Lookup de prestaciones por rol de MO para etiqueta real del encabezado
  const manoObraPrestById = {};
  sheetToObjects(ssCot, "ManoObra").forEach(m => {
    manoObraPrestById[String(m.id)] = parseFloat(m.prestaciones_pct) || 0;
  });

  sheet.setColumnWidth(1, 45);
  sheet.setColumnWidth(2, 285);
  sheet.setColumnWidth(3, 68);
  sheet.setColumnWidth(4, 72);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 130);

  let r = 1;

  // ─── LOGO + ENCABEZADO ────────────────────────────────────────────────────────
  sheet.setRowHeight(r, 55);
  sheet.getRange(r, 1, 1, 2).merge().setValue("").setBackground("#ffffff");
  sheet.getRange(r, 3, 1, 4).merge()
    .setValue(empresa || "").setFontSize(16).setFontWeight("bold").setFontColor("#000000")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#ffffff");
  if (logoId) {
    try {
      const img = sheet.insertImage(DriveApp.getFileById(logoId).getBlob(), 1, r);
      img.setWidth(160).setHeight(48);
    } catch(e) {}
  }
  r++;

  // Fila cliente + dirección (con wrap para que no se corte)
  const clienteTexto   = cot.cliente   ? "CLIENTE:    " + cot.cliente   : "";
  const direccionTexto = cot.direccion ? "DIRECCIÓN:  " + cot.direccion : "";
  sheet.setRowHeight(r, 22);
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
  sheet.setRowHeight(r, 14);
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
  sheet.setRowHeight(r, 20);
  sheet.getRange(r, 1, 1, NC)
    .setValues([["ÍTEM", "DESCRIPCIÓN", "UND", "CANT.", "VR. UNITARIO", "VR. TOTAL"]])
    .setBackground(NEGRO).setFontColor("#1a1a1a").setFontWeight("bold").setFontSize(8)
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

    const desperdicioPct    = parseFloat(item.desperdicio_pct)       || 0;
    const hmPct             = parseFloat(item.herramienta_menor_pct) || 0;
    const desperdicioFactor = 1 + desperdicioPct / 100;
    const subMoBase         = (item.mano_obra || []).reduce((s, it) => s + (parseFloat(it.valor_parcial) || 0), 0);
    const hmValor           = subMoBase * hmPct / 100;

    // Fila principal del APU
    sheet.setRowHeight(r, 18);
    sheet.getRange(r, 1, 1, NC)
      .setBackground(GRIS_OSC)
      .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 1).setValue(idx + 1)
      .setFontSize(9).setFontWeight("bold").setFontColor("#1a1a1a")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(r, 2).setValue((item.codigo_item ? "[" + item.codigo_item + "]  " : "") + (item.descripcion || ""))
      .setFontSize(9).setFontWeight("bold").setFontColor("#1a1a1a")
      .setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true);
    sheet.getRange(r, 3).setValue(item.unidad || "")
      .setFontSize(9).setFontColor("#1a1a1a").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(r, 4).setValue(cant)
      .setFontSize(9).setFontColor("#1a1a1a").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange(r, 5).setValue(precio).setNumberFormat(MONEY)
      .setFontSize(9).setFontColor("#1a1a1a").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange(r, 6).setValue(total).setNumberFormat(MONEY)
      .setFontSize(9).setFontWeight("bold").setFontColor("#1a1a1a")
      .setHorizontalAlignment("right").setVerticalAlignment("middle");
    r++;

    // Sub-secciones del APU
    SEC.forEach(sec => {
      const subitems = item[sec.key] || [];
      const renderHM = sec.key === "equipos" && hmPct > 0;
      if (!subitems.length && !renderHM) return;

      // Encabezado de sección — con sufijo informativo
      let labelSec = sec.label;
      if (sec.key === "materiales" && desperdicioPct > 0) {
        labelSec += "   (Incluye factor de desperdicio del " + (desperdicioPct % 1 === 0 ? desperdicioPct.toFixed(0) : desperdicioPct.toString()) + "%)";
      } else if (sec.key === "mano_obra") {
        labelSec += "   " + construirSufijoPrestaciones(item.mano_obra, manoObraPrestById);
      }
      sheet.setRowHeight(r, 12);
      sheet.getRange(r, 1, 1, NC).merge()
        .setValue("  " + labelSec)
        .setBackground(GRIS_CLR).setFontColor("#333333")
        .setFontWeight("bold").setFontSize(8).setFontStyle("italic")
        .setHorizontalAlignment("left").setVerticalAlignment("middle")
        .setBorder(true, true, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
      r++;

      // Ítems de la sección
      subitems.forEach((it, i) => {
        const bg = i % 2 === 0 ? "#ffffff" : "#fafafa";
        const cantTexto = it.cantidad != null ? it.cantidad : "";
        const mostrarUnidad = sec.key === "materiales";
        sheet.getRange(r, 1).setValue(resolverCodigoItem(it))
          .setFontSize(7).setFontColor("#666666")
          .setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBackground(bg)
          .setBorder(true, true, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        if (mostrarUnidad) {
          sheet.getRange(r, 2).setValue("      " + (it.descripcion_manual || "—"))
            .setFontSize(8).setFontColor("#222222").setHorizontalAlignment("left")
            .setVerticalAlignment("middle").setWrap(true).setBackground(bg)
            .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
          sheet.getRange(r, 3).setValue(it.unidad || "")
            .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("center")
            .setVerticalAlignment("middle").setBackground(bg)
            .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        } else {
          sheet.getRange(r, 2, 1, 2).merge().setValue("      " + (it.descripcion_manual || "—"))
            .setFontSize(8).setFontColor("#222222").setHorizontalAlignment("left")
            .setVerticalAlignment("middle").setWrap(true).setBackground(bg)
            .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        }
        sheet.getRange(r, 4).setValue(cantTexto)
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 5).setValue(parseFloat(it.precio_unitario) || 0).setNumberFormat(MONEY)
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        const vpBase     = parseFloat(it.valor_parcial) || 0;
        const vpMostrado = sec.key === "materiales" ? vpBase * desperdicioFactor : vpBase;
        sheet.getRange(r, 6).setValue(vpMostrado).setNumberFormat(MONEY)
          .setFontSize(8).setFontWeight("bold").setFontColor("#111111").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.setRowHeight(r, 13);
        r++;
      });

      // Línea automática Herramienta Menor (HM) al final de Equipos
      if (renderHM) {
        const bg = subitems.length % 2 === 0 ? "#ffffff" : "#fafafa";
        sheet.setRowHeight(r, 13);
        sheet.getRange(r, 1).setValue("HM")
          .setFontSize(7).setFontWeight("bold").setFontColor("#666666")
          .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, true, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 2, 1, 2).merge().setValue("      Herramienta Menor (% MO)")
          .setFontSize(8).setFontColor("#222222").setHorizontalAlignment("left")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 4).setValue(hmPct / 100).setNumberFormat("0.##%")
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 5).setValue(subMoBase).setNumberFormat(MONEY)
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 6).setValue(hmValor).setNumberFormat(MONEY)
          .setFontSize(8).setFontWeight("bold").setFontColor("#111111").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        r++;
      }

      // Subtotal de sección (con desperdicio en materiales / HM en equipos)
      let secSubtotal = subitems.reduce((s, it) => s + (parseFloat(it.valor_parcial) || 0), 0);
      if (sec.key === "materiales") secSubtotal *= desperdicioFactor;
      else if (renderHM)             secSubtotal += hmValor;
      sheet.setRowHeight(r, 15);
      sheet.getRange(r, 1, 1, 5).merge()
        .setValue("Subtotal " + sec.label)
        .setFontSize(8).setFontWeight("bold").setFontColor("#333333")
        .setHorizontalAlignment("right").setVerticalAlignment("middle")
        .setBackground(GRIS_CLR)
        .setBorder(true, true, true, null, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(r, 6).setValue(secSubtotal).setNumberFormat(MONEY)
        .setFontSize(8).setFontWeight("bold").setFontColor("#111111")
        .setHorizontalAlignment("right").setVerticalAlignment("middle")
        .setBackground(GRIS_CLR)
        .setBorder(true, null, true, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
      r++;
    });

    // Fila de subtotal del APU
    sheet.setRowHeight(r, 15);
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
  sheet.setRowHeight(r, 18);
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("TOTAL COSTOS DIRECTOS")
    .setFontSize(9).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(GRIS_OSC)
    .setBorder(true, true, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(r, 6).setValue(valorNeto).setNumberFormat(MONEY)
    .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
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
    sheet.setRowHeight(r, 15);
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
  sheet.setRowHeight(r, 20);
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("VALOR TOTAL OFERTA")
    .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(NEGRO)
    .setBorder(true, true, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(r, 6).setValue(total).setNumberFormat(MONEY)
    .setFontSize(12).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(NEGRO)
    .setBorder(true, null, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  r++;

  // ─── CONDICIONES COMERCIALES ──────────────────────────────────────────────────
  if (formaPago || plazoEntrega || validezOferta || noIncluye) {
    r += 1;
    sheet.setRowHeight(r, 18);
    sheet.getRange(r, 1, 1, NC).merge()
      .setValue("CONDICIONES COMERCIALES")
      .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
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
  r += 1;
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

    // Build unidad lookup from BD (APU_Items doesn't store unidad)
    const matUnidad  = {};
    sheetToObjects(ss, "Materiales").forEach(m => { matUnidad[String(m.id)] = m.unidad || ""; });
    const otroUnidad = {};
    sheetToObjects(ss, "Otros").forEach(o => { otroUnidad[String(o.id)] = o.unidad || ""; });

    allApuItems.forEach(ai => {
      if (!ai.unidad) {
        if      (ai.tipo === "MATERIAL")                         ai.unidad = matUnidad[String(ai.recurso_id)] || "";
        else if (ai.tipo === "EQUIPO" || ai.tipo === "MANO_OBRA") ai.unidad = "Día";
        else if (ai.tipo === "OTRO")                             ai.unidad = otroUnidad[String(ai.recurso_id)] || "";
      }
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

  // Adjuntar desperdicio_pct y herramienta_menor_pct por APU
  const apuSheet = ss.getSheetByName("APU");
  if (apuSheet) {
    const apuData  = apuSheet.getDataRange().getValues();
    if (apuData.length > 1) {
      const apuH    = apuData[0];
      const idxId   = apuH.indexOf("id");
      const idxDesp = apuH.indexOf("desperdicio_pct");
      const idxHm   = apuH.indexOf("herramienta_menor_pct");
      const apuPctById = {};
      for (let i = 1; i < apuData.length; i++) {
        const id = apuData[i][idxId];
        apuPctById[String(id)] = {
          desperdicio_pct:       idxDesp >= 0 ? (parseFloat(apuData[i][idxDesp]) || 0) : 0,
          herramienta_menor_pct: idxHm   >= 0 ? (parseFloat(apuData[i][idxHm])   || 0) : 0,
        };
      }
      cot.items.forEach(item => {
        const p = apuPctById[String(item.apu_id)] || { desperdicio_pct: 0, herramienta_menor_pct: 0 };
        item.desperdicio_pct       = p.desperdicio_pct;
        item.herramienta_menor_pct = p.herramienta_menor_pct;
      });
    }
  }

  return cot;
}

// Recalcula item_num de TODOS los ítems de una cotización como "capitulo.sufijo".
// Agrupa por capitulo_num en el orden en que aparecen las filas; dentro de cada
// grupo numera 1, 2, 3... Ej: cap "3" → 3.1, 3.2 · cap "4" → 4.1.
// Si un ítem no tiene capitulo_num, usa "1".
function renumerarCotizacion(ss, cotId) {
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const h        = data[0];
  const cIdx     = h.indexOf("cotizacion_id");
  const numIdx   = h.indexOf("item_num");
  const capIdx   = h.indexOf("capitulo_num");

  const contadorPorCapitulo = {}; // { "3": 2, "4": 1 }
  for (let i = 1; i < data.length; i++) {
    if (data[i][cIdx] != cotId) continue;
    let cap = capIdx >= 0 ? String(data[i][capIdx] || "").trim() : "";
    if (!cap) cap = "1";
    contadorPorCapitulo[cap] = (contadorPorCapitulo[cap] || 0) + 1;
    const nuevoNum = cap + "." + contadorPorCapitulo[cap];
    sheet.getRange(i + 1, numIdx + 1).setValue(nuevoNum);
  }
}

// Cambia el capitulo_num y/o capitulo_nombre de un ítem y renumera la cotización.
function asignarCapituloItem(itemId, capituloNum, capituloNombre) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  const h      = data[0];
  const capIdx = h.indexOf("capitulo_num");
  const nomIdx = h.indexOf("capitulo_nombre");
  const cIdx   = h.indexOf("cotizacion_id");

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const cotId = data[i][cIdx];
      if (capituloNum !== undefined && capIdx >= 0) {
        sheet.getRange(i + 1, capIdx + 1).setValue(String(capituloNum || "1").trim() || "1");
      }
      if (capituloNombre !== undefined && nomIdx >= 0) {
        sheet.getRange(i + 1, nomIdx + 1).setValue(String(capituloNombre || "").trim());
      }
      renumerarCotizacion(ss, cotId);
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─── AGREGAR APU A COTIZACIÓN ─────────────────────────────────────────────────

function agregarAPUaCotizacion(cotId, apuId, cantidad, capituloNum, capituloNombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const apu = getAPUCompleto(apuId);
  if (!apu) return { ok: false };

  const sheet  = ss.getSheetByName("Cotizacion_Items");
  const data   = sheet.getDataRange().getValues();
  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId  = lastId + 1;

  const cant      = parseFloat(cantidad) || 1;
  const precioAPU = parseFloat(apu.costo_neto) || 0;
  const valTotal  = cant * precioAPU;
  const cap       = String(capituloNum || "1").trim() || "1";
  const capNom    = String(capituloNombre || "").trim();

  // Orden de columnas: id, cotizacion_id, apu_id, item_num, descripcion,
  // unidad, cantidad, precio_apu, valor_total, capitulo_num, capitulo_nombre
  sheet.appendRow([
    newId, cotId, apuId, "",            // item_num se asigna en renumerarCotizacion
    apu.descripcion || apu.codigo_item || "",
    apu.unidad || "",
    cant, precioAPU, valTotal,
    cap, capNom
  ]);

  const datosCliente = heredarDatosClienteSiVacios(ss, cotId, apu);
  renumerarCotizacion(ss, cotId);
  recalcularCotizacion(ss, cotId);

  return {
    id:          newId,
    valor_total: valTotal,
    cliente:     datosCliente.cliente,
    direccion:   datosCliente.direccion,
    apu: {
      equipos:    apu.equipos    || [],
      materiales: apu.materiales || [],
      mano_obra:  apu.mano_obra  || [],
      otros:      apu.otros      || [],
    },
  };
}

// Copia cliente/dirección del APU a la cotización SOLO si la cotización los tiene
// vacíos. Devuelve los valores resultantes (heredados o los que ya tenía).
function heredarDatosClienteSiVacios(ss, cotId, apu) {
  const sheet = ss.getSheetByName("Cotizaciones");
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];
  const cCli  = h.indexOf("cliente");
  const cDir  = h.indexOf("direccion");

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == cotId) {
      let cliente   = cCli >= 0 ? String(data[i][cCli] || "").trim() : "";
      let direccion = cDir >= 0 ? String(data[i][cDir] || "").trim() : "";

      if (!cliente && apu.cliente) {
        cliente = String(apu.cliente).trim();
        if (cCli >= 0) sheet.getRange(i + 1, cCli + 1).setValue(cliente);
      }
      if (!direccion && apu.direccion) {
        direccion = String(apu.direccion).trim();
        if (cDir >= 0) sheet.getRange(i + 1, cDir + 1).setValue(direccion);
      }
      return { cliente: cliente, direccion: direccion };
    }
  }
  return { cliente: "", direccion: "" };
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
      const campos = ["cliente","direccion","administracion_pct","imprevistos_pct","utilidad_pct","iva_pct","aprobada","notas","forma_pago","plazo_entrega","validez_oferta","no_incluye","objeto"];
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

function enviarCotizacionEmail(cotId, emailDestino, tipo) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const { cot, render } = _prepararCotizacion(cotId, tipo);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Generar PDF
    const tmpName = "_cot_email_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    render(tmp, cot);
    SpreadsheetApp.flush();

    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=pdf&gid=" + tmp.getSheetId()
      + "&portrait=true&scale=4&size=letter"
      + "&gridlines=false&printtitle=false&sheetnames=false"
      + "&top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50";

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

function enviarYGuardarPDF(cotId, emailDestino, tipo) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const { cot, render } = _prepararCotizacion(cotId, tipo);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Generar PDF una sola vez
    const tmpName = "_cot_enviar_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    render(tmp, cot);
    SpreadsheetApp.flush();

    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=pdf&gid=" + tmp.getSheetId()
      + "&portrait=true&scale=4&size=letter"
      + "&gridlines=false&printtitle=false&sheetnames=false"
      + "&top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50";

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

// ─── GET COTIZACIÓN CLIENTE (liviana, escala a 300+ ítems) ───────────────────
// A diferencia de getCotizacionCompleta, NO carga APU_Items ni la BD de materiales.
// El documento del cliente no muestra desglose, así que no se necesitan sub-ítems.
function getCotizacionCliente(cotId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cotSheet = ss.getSheetByName("Cotizaciones");
  const cotData  = cotSheet.getDataRange().getValues();
  const cotH     = cotData[0];
  const cotRow   = cotData.slice(1).find(r => r[0] == cotId);
  if (!cotRow) return null;

  const cot = {};
  cotH.forEach((h, i) => cot[h] = cotRow[i]);

  // Normalizar % guardados como decimal (0.5 → 50)
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

  // Ordenar por número de ítem (orden numérico jerárquico)
  cot.items.sort((a, b) => compararItemNum(a.item_num, b.item_num));

  return cot;
}

// ─── ACTUALIZAR ÍTEM DE COTIZACIÓN (campos múltiples) ────────────────────────────
// Permite editar item_num, descripcion, unidad, cantidad y precio_apu.
// Recalcula valor_total (= cantidad * precio_apu) y los totales de la cotización.
function actualizarItemCotizacion(itemId, cambios) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const row   = [...data[i]];
      const cotId = row[h.indexOf("cotizacion_id")];

      ["item_num", "descripcion", "unidad"].forEach(campo => {
        if (cambios[campo] !== undefined) {
          const c = h.indexOf(campo);
          if (c >= 0) row[c] = cambios[campo];
        }
      });

      const cant   = parseFloat(cambios.cantidad   ?? row[h.indexOf("cantidad")])   || 0;
      const precio = parseFloat(cambios.precio_apu ?? row[h.indexOf("precio_apu")]) || 0;
      row[h.indexOf("cantidad")]    = cant;
      row[h.indexOf("precio_apu")]  = precio;
      row[h.indexOf("valor_total")] = cant * precio;

      sheet.getRange(i + 1, 1, 1, h.length).setValues([row]);
      recalcularCotizacion(ss, cotId);
      return { ok: true, valor_total: cant * precio };
    }
  }
  return { ok: false };
}

// ─── AGREGAR LÍNEA MANUAL (sin APU) ──────────────────────────────────────────
function agregarLineaManual(cotId, datos) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName("Cotizacion_Items");
  const data   = sheet.getDataRange().getValues();
  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId  = lastId + 1;

  const cant   = parseFloat(datos.cantidad)   || 0;
  const precio = parseFloat(datos.precio_apu) || 0;

  // Orden de columnas: id, cotizacion_id, apu_id, item_num, descripcion, unidad, cantidad, precio_apu, valor_total
  sheet.appendRow([
    newId, cotId, "",
    datos.item_num    || "",
    datos.descripcion || "",
    datos.unidad      || "",
    cant, precio, cant * precio
  ]);

  recalcularCotizacion(ss, cotId);
  return { id: newId, valor_total: cant * precio };
}

// ─── AGREGAR VARIAS LÍNEAS DE UN GOLPE (escala a 300+ ítems) ─────────────────
// lineas: [{ item_num, descripcion, unidad, cantidad, precio_apu }, ...]
// Una sola escritura a la hoja + un solo recálculo.
function agregarVariasLineas(cotId, lineas) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  let nextId  = (data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0) + 1;

  const filas = (lineas || []).map(l => {
    const cant   = parseFloat(l.cantidad)   || 0;
    const precio = parseFloat(l.precio_apu) || 0;
    return [
      nextId++, cotId, "",
      l.item_num    || "",
      l.descripcion || "",
      l.unidad      || "",
      cant, precio, cant * precio
    ];
  });
  if (!filas.length) return { ok: true, count: 0 };

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, filas.length, filas[0].length).setValues(filas);
  recalcularCotizacion(ss, cotId);
  return { ok: true, count: filas.length };
}

// ─── RENDER DOCUMENTO DEL CLIENTE (tabla limpia, sin desglose) ───────────────
// Diseñado para 300+ ítems: la tabla de ítems se escribe con UN solo setValues
// y los formatos se aplican por rangos completos, no celda por celda.
function llenarHojaCotizacionCliente(sheet, cot) {
  const items   = (cot.items || []).slice().sort((a, b) => compararItemNum(a.item_num, b.item_num));
  const valorNeto = parseFloat(cot.valor_neto) || 0;
  const admPct  = parseFloat(cot.administracion_pct) || 0;
  const impPct  = parseFloat(cot.imprevistos_pct)    || 0;
  const utilPct = parseFloat(cot.utilidad_pct)       || 0;
  const ivaPct  = parseFloat(cot.iva_pct)            || 0;

  const cfg       = getConfig();
  const empresa   = (cfg["empresa"]          || "").trim();
  const remitente = (cfg["nombre_remitente"] || "").trim();
  const logoId    = (cfg["logo_id"]          || "").trim();
  const firmaId   = (cfg["firma_id"]         || "").trim();

  const formaPago     = String(cot.forma_pago     || "").trim();
  const plazoEntrega  = String(cot.plazo_entrega  || "").trim();
  const validezOferta = String(cot.validez_oferta || "").trim();
  const noIncluye     = String(cot.no_incluye     || "").trim();

  const MONEY = '"$"#,##0';
  const NEGRO = "#a8a8a8", GRIS_OSC = "#c8c8c8", BORDE = "#dddddd";
  const NC = 6;

  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 55);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 130);

  let r = 1;

  // Encabezado: logo + empresa
  sheet.setRowHeight(r, 55);
  sheet.getRange(r, 1, 1, 2).merge().setValue("").setBackground("#ffffff");
  sheet.getRange(r, 3, 1, 4).merge()
    .setValue(empresa || "").setFontSize(16).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#ffffff");
  if (logoId) {
    try { sheet.insertImage(DriveApp.getFileById(logoId).getBlob(), 1, r).setWidth(160).setHeight(48); } catch(e) {}
  }
  r++;

  // Cliente + dirección
  sheet.setRowHeight(r, 22);
  sheet.getRange(r, 1, 1, 3).merge().setValue(cot.cliente ? "CLIENTE:    " + cot.cliente : "")
    .setFontSize(9).setFontColor("#222222").setFontWeight("bold")
    .setVerticalAlignment("middle").setWrap(true).setBackground("#f5f5f5");
  sheet.getRange(r, 4, 1, 3).merge().setValue(cot.direccion ? "DIRECCIÓN:  " + cot.direccion : "")
    .setFontSize(9).setFontColor("#333333").setVerticalAlignment("middle").setWrap(true).setBackground("#f5f5f5");
  r++;

  // N° oferta + fecha
  sheet.setRowHeight(r, 14);
  const infoOferta = [
    cot.numero_oferta ? "N° Oferta: " + cot.numero_oferta : null,
    cot.fecha         ? "Fecha: "     + cot.fecha         : null,
  ].filter(Boolean).join("     ");
  sheet.getRange(r, 1, 1, NC).merge().setValue(infoOferta)
    .setFontSize(9).setFontColor("#555555").setHorizontalAlignment("right").setVerticalAlignment("middle");
  r++;

  // Encabezado de tabla
  sheet.setRowHeight(r, 20);
  sheet.getRange(r, 1, 1, NC)
    .setValues([["ÍTEM", "DESCRIPCIÓN", "UND", "CANT.", "VR. UNITARIO", "VR. TOTAL"]])
    .setBackground(NEGRO).setFontColor("#1a1a1a").setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  r++;

  // ── ÍTEMS: una sola escritura en bloque ──
  const tablaInicio = r;
  if (items.length) {
    const matriz = items.map(it => {
      const cant   = parseFloat(it.cantidad)   || 0;
      const precio = parseFloat(it.precio_apu) || 0;
      return [
        it.item_num || "",
        it.descripcion || "—",
        it.unidad || "",
        cant,
        precio,
        cant * precio,
      ];
    });
    const rango = sheet.getRange(tablaInicio, 1, matriz.length, NC);
    rango.setValues(matriz);
    rango.setFontSize(9).setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(tablaInicio, 1, matriz.length, 1).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(tablaInicio, 2, matriz.length, 1).setHorizontalAlignment("left").setWrap(true);
    sheet.getRange(tablaInicio, 3, matriz.length, 1).setHorizontalAlignment("center");
    sheet.getRange(tablaInicio, 4, matriz.length, 1).setHorizontalAlignment("right");
    sheet.getRange(tablaInicio, 5, matriz.length, 2).setHorizontalAlignment("right").setNumberFormat(MONEY);
    sheet.getRange(tablaInicio, 6, matriz.length, 1).setFontWeight("bold");
    r += matriz.length;
  } else {
    sheet.getRange(r, 1, 1, NC).merge().setValue("Sin ítems")
      .setFontSize(9).setFontColor("#888888").setHorizontalAlignment("center");
    r++;
  }

  // ── TOTALES (Subtotal + AIU + IVA + Total) ──
  const admVal  = Math.round(valorNeto * admPct  / 100);
  const impVal  = Math.round(valorNeto * impPct  / 100);
  const utilVal = Math.round(valorNeto * utilPct / 100);
  const sinIVA  = valorNeto + admVal + impVal + utilVal;
  const ivaVal  = Math.round(sinIVA * ivaPct / 100);
  const total   = sinIVA + ivaVal;

  const filasTot = [["SUBTOTAL", valorNeto]];
  if (admPct  > 0) filasTot.push(["ADMINISTRACIÓN (" + admPct  + "%)", admVal]);
  if (impPct  > 0) filasTot.push(["IMPREVISTOS ("    + impPct  + "%)", impVal]);
  if (utilPct > 0) filasTot.push(["UTILIDAD ("       + utilPct + "%)", utilVal]);
  if (ivaPct  > 0) filasTot.push(["IVA ("            + ivaPct  + "%)", ivaVal]);

  filasTot.forEach(([label, val]) => {
    sheet.setRowHeight(r, 16);
    sheet.getRange(r, 1, 1, 5).merge().setValue(label)
      .setFontSize(9).setFontColor("#333333").setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBorder(true, true, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 6).setValue(val).setNumberFormat(MONEY)
      .setFontSize(9).setFontColor("#333333").setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBorder(true, null, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  sheet.setRowHeight(r, 22);
  sheet.getRange(r, 1, 1, 5).merge().setValue("VALOR TOTAL OFERTA")
    .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(NEGRO)
    .setBorder(true, true, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(r, 6).setValue(total).setNumberFormat(MONEY)
    .setFontSize(12).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(NEGRO)
    .setBorder(true, null, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  r++;

  // ── Condiciones comerciales ──
  if (formaPago || plazoEntrega || validezOferta || noIncluye) {
    r += 1;
    sheet.setRowHeight(r, 18);
    sheet.getRange(r, 1, 1, NC).merge().setValue("CONDICIONES COMERCIALES")
      .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a").setBackground(GRIS_OSC)
      .setVerticalAlignment("middle");
    r++;
    [
      formaPago     ? ["FORMA DE PAGO",        formaPago]     : null,
      plazoEntrega  ? ["PLAZO DE ENTREGA",     plazoEntrega]  : null,
      validezOferta ? ["VALIDEZ DE LA OFERTA", validezOferta] : null,
    ].filter(Boolean).forEach(([label, val]) => {
      sheet.setRowHeight(r, 20);
      sheet.getRange(r, 1, 1, 2).merge().setValue(label)
        .setFontSize(9).setFontWeight("bold").setFontColor("#333333").setBackground("#f5f5f5").setVerticalAlignment("middle");
      sheet.getRange(r, 3, 1, 4).merge().setValue(val)
        .setFontSize(9).setFontColor("#333333").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("middle");
      r++;
    });
    if (noIncluye) {
      r++;
      sheet.getRange(r, 1, 1, NC).merge().setValue("NUESTRA OFERTA NO INCLUYE:")
        .setFontSize(9).setFontWeight("bold").setFontColor("#333333");
      r++;
      sheet.setRowHeight(r, 60);
      sheet.getRange(r, 1, 1, NC).merge().setValue(noIncluye)
        .setFontSize(9).setFontColor("#555555").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top");
      r++;
    }
  }

  // ── Firma ──
  r += 1;
  if (firmaId) {
    try {
      sheet.setRowHeight(r, 65);
      sheet.insertImage(DriveApp.getFileById(firmaId).getBlob(), 1, r).setWidth(200).setHeight(58);
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
    sheet.getRange(r, 1, 1, NC).merge().setValue(firmaLinea)
      .setFontSize(9).setFontWeight("bold").setVerticalAlignment("top")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
}

// ─── HELPER: ELEGIR DATOS + RENDERER SEGÚN TIPO ────────────────────────────────
// Elige datos + renderer según tipo ("cliente" por defecto).
function _prepararCotizacion(cotId, tipo) {
  if (tipo === "interna") {
    return { cot: getCotizacionCompleta(cotId), render: llenarHojaCotizacion };
  }
  return { cot: getCotizacionCliente(cotId), render: llenarHojaCotizacionCliente };
}


// Genera un .xlsx del documento del cliente desde la fuente única y lo guarda en Drive.
// Reusa el render limpio llenarHojaCotizacionCliente (sin desglose interno).
function exportarClienteXlsxV2(cotId) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const cot = getCotizacionCliente(cotId);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Hoja temporal con el documento renderizado
    const tmpName = "_cot_xlsx_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    llenarHojaCotizacionCliente(tmp, cot);
    SpreadsheetApp.flush();

    // Exportar SOLO esa hoja como xlsx
    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=xlsx&gid=" + tmp.getSheetId();
    const blob = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
    }).getBlob().setName("Cotizacion_" + (cot.numero_oferta || cotId) + ".xlsx");

    ss.deleteSheet(tmp);

    // Guardar en la carpeta de Drive configurada (o raíz si no hay)
    const cfg       = getConfig();
    const carpetaId = (cfg["carpeta_cotizaciones_cliente"] || "").trim();
    let archivo;
    if (carpetaId) {
      archivo = DriveApp.getFolderById(carpetaId).createFile(blob);
    } else {
      archivo = DriveApp.createFile(blob);
    }
    return { ok: true, nombre: archivo.getName(), url: archivo.getUrl() };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
