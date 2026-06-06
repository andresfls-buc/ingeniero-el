// ─── COTIZACIÓN CLIENTE — CRUD ────────────────────────────────────────────────
// Hoja: Cotizacion_Cliente_Items
// Columnas: id, cotizacion_id, item_num, descripcion, unidad, cantidad, precio_unitario, valor_total

function _clienteSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Cotizacion_Cliente_Items");
  // Auto-crear la hoja si no existe — sin migración manual ni alertas de UI,
  // para que la página de Cotizaciones Clientes funcione de inmediato.
  if (!sheet) {
    sheet = ss.insertSheet("Cotizacion_Cliente_Items");
    sheet.appendRow(["id", "cotizacion_id", "item_num", "descripcion", "unidad", "cantidad", "precio_unitario", "valor_total"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Devuelve { cot: {...}, items: [...], cfg: {...} }
function getCotizacionClienteDoc(cotId) {
  // DIAG: devolvemos SIEMPRE un string JSON. Un string nunca falla la
  // serialización de google.script.run, así vemos el contenido crudo real.
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Cabecera de la cotización (cliente, fecha, numero_oferta, AIU, condiciones)
    const cotSheet = ss.getSheetByName("Cotizaciones");
    const cotData  = cotSheet.getDataRange().getValues();
    const cotH     = cotData[0];
    const cotRow   = cotData.slice(1).find(r => r[0] == cotId);
    if (!cotRow) {
      return JSON.stringify({
        __diag: true,
        encontrada: false,
        cotIdRecibido: cotId,
        tipoCotId: typeof cotId,
        spreadsheetId: ss.getId(),
        spreadsheetNombre: ss.getName(),
        sheetCotizaciones: cotSheet ? cotSheet.getName() : "NO EXISTE",
        totalFilas: cotData.length - 1,
        idsEnHoja: cotData.slice(1).map(r => r[0]),
        headers: cotH,
      });
    }
    const cot = {};
    cotH.forEach((h, i) => cot[h] = cotRow[i]);
    ["administracion_pct", "imprevistos_pct", "utilidad_pct", "iva_pct"].forEach(k => {
      const v = parseFloat(cot[k]) || 0;
      cot[k] = (v > 0 && v < 1) ? Math.round(v * 100) : v;
    });

    // Ítems del documento cliente
    const sheet = _clienteSheet();
    const data  = sheet.getDataRange().getValues();
    let items = [];
    if (data.length > 1) {
      const h = data[0];
      items = data.slice(1)
        .filter(r => r[h.indexOf("cotizacion_id")] == cotId)
        .map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
      items.sort((a, b) => compararItemNum(a.item_num, b.item_num));
    }

    // Config de empresa (para header del documento)
    const cfg = getConfig();

    return JSON.stringify({ encontrada: true, cot, items, cfg });
  } catch (e) {
    return JSON.stringify({ __diag: true, error: String(e && e.message || e), stack: String(e && e.stack || "") });
  }
}

// Agrega una fila vacía. Devuelve { id, valor_total: 0 }
function agregarFilaClienteDoc(cotId) {
  const sheet   = _clienteSheet();
  const data    = sheet.getDataRange().getValues();
  const lastId  = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId = lastId + 1;
  sheet.appendRow([newId, cotId, "", "", "", 1, 0, 0]);
  return { id: newId, valor_total: 0 };
}

// Puebla Cotizacion_Cliente_Items desde los APUs de la cotización interna (solo si está vacío).
// Para cada ítem interno con apu_id, crea una fila con: actividad→descripcion, unidad, costo_neto→precio_unitario.
function inicializarClienteDesdeAPUs(cotId) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = _clienteSheet();
  const data    = sheet.getDataRange().getValues();
  const h       = data[0];

  // Solo inicializar si no hay filas para este cotId
  const yaExisten = data.length > 1 &&
    data.slice(1).some(r => r[h.indexOf("cotizacion_id")] == cotId);
  if (yaExisten) return { ok: true, initialized: false, items: [] };

  // Leer ítems internos de la cotización
  const cotItemsData = ss.getSheetByName("Cotizacion_Items").getDataRange().getValues();
  const cih          = cotItemsData[0];
  const internos     = cotItemsData.slice(1)
    .filter(r => r[cih.indexOf("cotizacion_id")] == cotId && r[cih.indexOf("apu_id")]);

  if (internos.length === 0) return { ok: true, initialized: false, items: [] };

  // Leer APUs para obtener actividad, unidad, costo_neto
  const apuData = ss.getSheetByName("APU").getDataRange().getValues();
  const ah      = apuData[0];
  const apuById = {};
  apuData.slice(1).forEach(r => {
    apuById[String(r[ah.indexOf("id")])] = r;
  });

  // Calcular el último id existente en Cotizacion_Cliente_Items
  const allData = sheet.getDataRange().getValues();
  const lastId  = allData.length > 1
    ? Math.max(...allData.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;

  let nextId  = lastId + 1;
  const created = [];

  internos.forEach(interno => {
    const apuId  = String(interno[cih.indexOf("apu_id")]);
    const apuRow = apuById[apuId];
    if (!apuRow) return;

    const itemNum   = interno[cih.indexOf("item_num")]  || "";
    const cantidad  = parseFloat(interno[cih.indexOf("cantidad")]) || 1;
    const desc      = apuRow[ah.indexOf("actividad")]   || apuRow[ah.indexOf("descripcion")] || "";
    const unidad    = apuRow[ah.indexOf("unidad")]      || "";
    const precio    = Math.round(parseFloat(apuRow[ah.indexOf("costo_neto")]) || 0);
    const total     = Math.round(cantidad * precio);

    sheet.appendRow([nextId, cotId, itemNum, desc, unidad, cantidad, precio, total]);
    created.push({ id: nextId, item_num: itemNum, descripcion: desc, unidad, cantidad, precio_unitario: precio, valor_total: total });
    nextId++;
  });

  return { ok: true, initialized: true, items: created };
}

// Actualiza campos de una fila. Recalcula valor_total.
function actualizarFilaClienteDoc(itemId, cambios) {
  const sheet = _clienteSheet();
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const row = [...data[i]];
      ["item_num", "descripcion", "unidad"].forEach(campo => {
        if (cambios[campo] !== undefined) row[h.indexOf(campo)] = cambios[campo];
      });
      const cant   = parseFloat(cambios.cantidad       ?? row[h.indexOf("cantidad")])       || 0;
      const precio = parseFloat(cambios.precio_unitario ?? row[h.indexOf("precio_unitario")]) || 0;
      row[h.indexOf("cantidad")]        = cant;
      row[h.indexOf("precio_unitario")] = precio;
      row[h.indexOf("valor_total")]     = Math.round(cant * precio);
      sheet.getRange(i + 1, 1, 1, h.length).setValues([row]);
      return { ok: true, valor_total: Math.round(cant * precio) };
    }
  }
  return { ok: false };
}

// Elimina una fila por id.
function eliminarFilaClienteDoc(itemId) {
  const sheet = _clienteSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

// Genera un .xlsx formateado del documento cliente y lo guarda en Drive.
// Devuelve { ok, url, nombre } o { ok: false, error }
function exportarCotizacionClienteXlsx(cotId) {
  try {
    const doc = getCotizacionClienteDoc(cotId);
    if (!doc) return { ok: false, error: "Cotización no encontrada" };
    const { cot, items, cfg } = doc;

    const folderId = (cfg["carpeta_cotizaciones_cliente"] || "").trim();
    if (!folderId) return { ok: false, error: "Configura 'carpeta_cotizaciones_cliente' en la hoja Configuracion." };

    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const tmpName = "_cot_cli_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);

    // ── LLENAR HOJA ──
    const empresa = cfg["empresa"]           || "";
    const nit     = cfg["nit"]               || "";
    const tel     = cfg["telefono"]          || "";
    const email   = cfg["email_empresa"]     || "";
    const nombre  = cfg["nombre_remitente"]  || "";
    const cargo   = cfg["cargo"]             || "";
    const tp      = cfg["tarjeta_profesional"] || "";

    const fecha = cot.fecha
      ? Utilities.formatDate(new Date(cot.fecha), Session.getScriptTimeZone(), "dd 'de' MMMM 'de' yyyy")
      : "";

    let row = 1;

    // Membrete
    tmp.getRange(row, 1, 1, 6).merge().setValue(empresa).setFontSize(14).setFontWeight("bold");
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue([nit ? "NIT: " + nit : "", tel, email].filter(Boolean).join("  ·  "))
      .setFontSize(10).setFontColor("#555555");
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue("OFERTA DE PRECIOS No. " + (cot.numero_oferta || ""))
      .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("right");
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue("Bogotá, " + fecha)
      .setHorizontalAlignment("right").setFontSize(10).setFontColor("#555555");
    row += 2;

    // Destinatario
    tmp.getRange(row, 1, 1, 6).merge().setValue("Señores: " + (cot.cliente || "")).setFontWeight("bold");
    row++;
    if (cot.direccion) {
      tmp.getRange(row, 1, 1, 6).merge().setValue("Dirección: " + cot.direccion);
      row++;
    }
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue("Por medio de la presente nos permitimos presentar nuestra oferta de precios para la ejecución de las siguientes actividades:")
      .setWrap(true);
    row += 2;

    // Encabezado tabla
    const headers = ["ÍTEM", "DESCRIPCIÓN", "UND", "CANT", "VR. UNITARIO", "VR. TOTAL"];
    const hRange  = tmp.getRange(row, 1, 1, 6);
    hRange.setValues([headers])
      .setBackground("#1a237e").setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    row++;

    // Filas de ítems
    const FCOP = v => "$" + Math.round(parseFloat(v) || 0).toLocaleString("es-CO");
    items.forEach(item => {
      const cant  = parseFloat(item.cantidad)        || 0;
      const precio= parseFloat(item.precio_unitario) || 0;
      const total = Math.round(cant * precio);
      tmp.getRange(row, 1, 1, 6).setValues([[
        item.item_num || "",
        item.descripcion || "",
        item.unidad || "",
        cant || "",
        FCOP(precio),
        FCOP(total),
      ]]);
      row++;
    });
    row++;

    // Totales
    const cd   = items.reduce((s, i) => s + Math.round((parseFloat(i.cantidad) || 0) * (parseFloat(i.precio_unitario) || 0)), 0);
    const admP = parseFloat(cot.administracion_pct) || 0;
    const impP = parseFloat(cot.imprevistos_pct)    || 0;
    const utlP = parseFloat(cot.utilidad_pct)       || 0;
    const ivaP = parseFloat(cot.iva_pct)            || 0;
    const adm  = Math.round(cd * admP / 100);
    const imp  = Math.round(cd * impP / 100);
    const utl  = Math.round(cd * utlP / 100);
    const sin  = cd + adm + imp + utl;
    const iva  = Math.round(sin * ivaP / 100);
    const tot  = sin + iva;

    [
      ["COSTO DIRECTO",                         FCOP(cd)],
      ["ADMINISTRACIÓN (" + admP + "%)",         FCOP(adm)],
      ["IMPREVISTOS (" + impP + "%)",            FCOP(imp)],
      ["UTILIDAD (" + utlP + "%)",               FCOP(utl)],
      ["SUBTOTAL SIN IVA",                       FCOP(sin)],
      ["IVA (" + ivaP + "%)",                    FCOP(iva)],
      ["VALOR TOTAL OFERTA",                     FCOP(tot)],
    ].forEach(([label, val], idx) => {
      const isTotal = idx === 6;
      tmp.getRange(row, 4, 1, 2).merge().setValue(label)
        .setHorizontalAlignment("right")
        .setFontWeight(isTotal ? "bold" : "normal")
        .setBackground(isTotal ? "#e8eaf6" : null);
      tmp.getRange(row, 6).setValue(val)
        .setHorizontalAlignment("right")
        .setFontWeight(isTotal ? "bold" : "normal")
        .setBackground(isTotal ? "#e8eaf6" : null);
      row++;
    });
    row++;

    // Condiciones comerciales
    [
      ["FORMA DE PAGO",    cot.forma_pago],
      ["PLAZO DE ENTREGA", cot.plazo_entrega],
      ["VALIDEZ OFERTA",   cot.validez_oferta],
      ["NO INCLUYE",       cot.no_incluye],
    ].filter(([, v]) => v).forEach(([label, val]) => {
      tmp.getRange(row, 1, 1, 6).merge().setValue(label + ":  " + val).setWrap(true);
      row++;
    });
    row += 2;

    // Cierre
    tmp.getRange(row, 1, 1, 6).merge().setValue("Quedamos atentos a sus consultas.  Cordialmente,");
    row += 2;
    tmp.getRange(row, 1, 1, 6).merge().setValue(nombre).setFontWeight("bold");
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue([cargo, tp ? "T.P. No. " + tp : ""].filter(Boolean).join("  ·  "));
    row++;
    tmp.getRange(row, 1, 1, 6).merge().setValue(empresa);

    // Formato columnas
    tmp.setColumnWidth(1, 55);
    tmp.setColumnWidth(2, 280);
    tmp.setColumnWidth(3, 55);
    tmp.setColumnWidth(4, 65);
    tmp.setColumnWidth(5, 110);
    tmp.setColumnWidth(6, 110);
    tmp.setRowHeights(1, tmp.getLastRow(), 22);

    SpreadsheetApp.flush();

    // Exportar como .xlsx y guardar en Drive
    const xlsxUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=xlsx&gid=" + tmp.getSheetId();
    const blob = UrlFetchApp.fetch(xlsxUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
    }).getBlob();

    const nombre_archivo = "COT-" + (cot.numero_oferta || cotId) + "_" + (cot.cliente || "Cliente").replace(/\s+/g, "_") + "_Cliente.xlsx";
    blob.setName(nombre_archivo);

    ss.deleteSheet(tmp);

    const folder = DriveApp.getFolderById(folderId);
    const file   = folder.createFile(blob);
    return { ok: true, url: file.getUrl(), nombre: nombre_archivo };

  } catch(e) {
    try {
      const bad = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_cot_cli_tmp_");
      if (bad) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(bad);
    } catch(_) {}
    return { ok: false, error: e.message };
  }
}
