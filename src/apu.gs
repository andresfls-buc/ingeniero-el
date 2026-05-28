// ─── DIAGNÓSTICO ──────────────────────────────────────────────────────────────

function diagnosticarBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { error: "getActiveSpreadsheet() retornó null — el script no está vinculado a ningún spreadsheet." };
  const hojas        = ss.getSheets().map(s => s.getName());
  const equiposSheet = ss.getSheetByName("Equipos");
  const equiposFilas = equiposSheet ? equiposSheet.getLastRow() : -1;
  const equiposCols  = equiposSheet ? equiposSheet.getLastColumn() : -1;
  const equiposParsed = equiposSheet ? sheetToObjects(ss, "Equipos").length : -1;
  return {
    spreadsheetNombre: ss.getName(),
    spreadsheetId:     ss.getId(),
    hojas:             hojas,
    equiposFilas:      equiposFilas,
    equiposCols:       equiposCols,
    equiposParsed:     equiposParsed,
  };
}

// ─── CARGA DE BD (por tipo, para evitar límite de serialización) ──────────────

function cargarEquipos() {
  const items = sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "Equipos");
  return items.map(e => ({
    id:         e.id,
    nombre:     e.nombre,
    tarifa_dia: e.tarifa_dia,
    partida:    e.partida || "",
  }));
}

function cargarMateriales() {
  const items = sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "Materiales");
  return items.map(m => ({
    id:             m.id,
    codigo:         m.codigo,
    categoria:      m.categoria,
    nombre:         m.nombre,
    unidad:         m.unidad,
    precio_sin_iva: m.precio_sin_iva,
  }));
}

function cargarManoObra() {
  return sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "ManoObra");
}

function cargarCategoriasMateriales() {
  const items = sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "Materiales");
  const cats = [...new Set(items.map(m => String(m.categoria || "")).filter(Boolean))].sort();
  return cats;
}

function buscarMateriales(query, categoria) {
  const items = sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "Materiales");
  const q   = String(query    || "").toLowerCase().trim();
  const cat = String(categoria || "").trim();
  return items.filter(m => {
    const matchCat = !cat || m.categoria === cat;
    const matchQ   = !q
      || String(m.nombre  || "").toLowerCase().includes(q)
      || String(m.codigo  || "").toLowerCase().includes(q);
    return matchCat && matchQ;
  }).slice(0, 50);
}

// ─── GENERAR PDF DEL APU ─────────────────────────────────────────────────────

function _generarBlobPDFApu(apuId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const apu = getAPUCompleto(apuId);
  if (!apu) return null;

  const tmpName = "_apu_pdf_tmp_";
  let tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  _llenarHojaAPU(tmp, apu, ss);
  SpreadsheetApp.flush();

  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId()
    + "/export?format=xlsx&gid=" + tmp.getSheetId();

  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getBlob();

  ss.deleteSheet(tmp);
  return { blob, apu };
}

function listarArchivosDrive() {
  const cfg = getConfig();
  const resultado = [];
  const carpetas = [
    { key: "carpeta_apus",         tipo: "APU" },
    { key: "carpeta_cotizaciones", tipo: "COT" },
  ];
  carpetas.forEach(({ key, tipo }) => {
    const id = (cfg[key] || "").trim();
    if (!id) return;
    try {
      _recolectarArchivos(DriveApp.getFolderById(id), tipo, resultado);
    } catch(e) {}
  });
  resultado.sort((a, b) => (b.fecha > a.fecha ? 1 : b.fecha < a.fecha ? -1 : 0));
  return resultado;
}

function _recolectarArchivos(folder, tipo, resultado) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    resultado.push({
      nombre: f.getName(),
      url:    f.getUrl(),
      fecha:  Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
      tipo,
    });
  }
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    _recolectarArchivos(subs.next(), tipo, resultado);
  }
}

function exportarAPUaDrive(apuId) {
  try {
    const result = _generarBlobPDFApu(apuId);
    if (!result) return { ok: false, error: "APU no encontrado" };
    const { blob, apu } = result;

    const cfg      = getConfig();
    const folderId = (cfg["carpeta_apus"] || "").trim();
    if (!folderId) return { ok: false, error: "Configura 'carpeta_apus' en la hoja Configuracion con el ID de tu carpeta de Drive." };

    const folder   = DriveApp.getFolderById(folderId);
    const slug     = String(apu.cliente || apu.descripcion || "apu")
      .normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/[^a-zA-Z0-9 ]/g, "").trim().replace(/\s+/g, "_");
    const codigo   = String(apu.codigo_item || "APU").replace(/[^a-zA-Z0-9_-]/g, "_");
    const fileName = codigo + "_" + slug + ".xlsx";

    // Reemplazar archivo anterior con el mismo nombre si existe
    const iter = folder.getFilesByName(fileName);
    while (iter.hasNext()) iter.next().setTrashed(true);

    blob.setName(fileName);
    folder.createFile(blob);
    return { ok: true, fileName };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function descargarAPUBase64(apuId) {
  try {
    const result = _generarBlobPDFApu(apuId);
    if (!result) return { ok: false, error: "APU no encontrado" };
    const { blob, apu } = result;

    const slug     = String(apu.cliente || apu.descripcion || "apu")
      .normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/[^a-zA-Z0-9 ]/g, "").trim().replace(/\s+/g, "_");
    const codigo   = String(apu.codigo_item || "APU").replace(/[^a-zA-Z0-9_-]/g, "_");
    const fileName = codigo + "_" + slug + ".xlsx";

    return { ok: true, base64: Utilities.base64Encode(blob.getBytes()), fileName };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function _llenarHojaAPU(sheet, apu, ss) {
  const MONEY    = '"$"#,##0';
  const NEGRO    = "#a8a8a8";
  const GRIS_OSC = "#c8c8c8";
  const GRIS_CLR = "#f5f5f5";
  const BORDE    = "#dddddd";
  const cfg      = getConfig();
  const empresa = (cfg["empresa"] || "").trim();
  const logoId  = (cfg["logo_id"] || "").trim();

  // Lookup de códigos de materiales (id → codigo)
  const materialesCodigoById = {};
  sheetToObjects(ss, "Materiales").forEach(m => {
    materialesCodigoById[String(m.id)] = m.codigo || "";
  });

  // Lookup de prestaciones por rol de MO (id → pct) para etiqueta real del encabezado
  const manoObraPrestById = {};
  sheetToObjects(ss, "ManoObra").forEach(m => {
    manoObraPrestById[String(m.id)] = parseFloat(m.prestaciones_pct) || 0;
  });

  const desperdicioPct   = parseFloat(apu.desperdicio_pct)        || 0;
  const hmPct            = parseFloat(apu.herramienta_menor_pct)  || 0;
  const desperdicioFactor= 1 + desperdicioPct / 100;
  const subtotalMOBase   = parseFloat(apu.subtotal_mano_obra)     || 0;
  const hmValor          = subtotalMOBase * hmPct / 100;

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 255);
  sheet.setColumnWidth(3, 62);
  sheet.setColumnWidth(4, 68);
  sheet.setColumnWidth(5, 68);
  sheet.setColumnWidth(6, 118);
  sheet.setColumnWidth(7, 118);

  let r = 1;

  // Fila 1: Logo + Título
  sheet.setRowHeight(r, 85);
  sheet.getRange(r, 1, 1, 2).merge().setValue("").setBackground("#ffffff");
  const titulo = ["APU", apu.codigo_item, empresa].filter(Boolean).join("  —  ");
  sheet.getRange(r, 3, 1, 5).merge()
    .setValue(titulo)
    .setFontSize(13).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#ffffff");
  if (logoId) {
    try {
      const img = sheet.insertImage(DriveApp.getFileById(logoId).getBlob(), 1, r);
      img.setWidth(215).setHeight(78);
    } catch(e) {}
  }
  r++;

  // Separador
  sheet.setRowHeight(r, 4);
  sheet.getRange(r, 1, 1, 7).merge().setValue("")
    .setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBackground("#ffffff");
  r++;

  // Datos generales — fila 1: cliente + código; fila 2: unidad + fecha
  const datosGen = [
    ["CLIENTE / OBRA", apu.cliente  || "—", "CÓDIGO", apu.codigo_item || "—"],
    ["UNIDAD",         apu.unidad   || "—", "FECHA",  apu.fecha       || "—"],
  ];
  datosGen.forEach((fila, idx) => {
    const isCliente = idx === 0;
    sheet.setRowHeight(r, isCliente ? 44 : 22);
    sheet.getRange(r, 1, 1, 2).merge().setValue(fila[0])
      .setFontSize(8).setFontWeight("bold").setFontColor("#555555")
      .setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBackground("#f5f5f5");
    const valorCell = sheet.getRange(r, 3, 1, 2).merge().setValue(fila[1])
      .setFontSize(9).setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBackground("#ffffff");
    if (isCliente) valorCell.setWrap(true).setVerticalAlignment("top");
    sheet.getRange(r, 5).setValue(fila[2])
      .setFontSize(8).setFontWeight("bold").setFontColor("#555555")
      .setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBackground("#f5f5f5");
    sheet.getRange(r, 6, 1, 2).merge().setValue(fila[3])
      .setFontSize(9).setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBackground("#ffffff");
    sheet.getRange(r, 1, 1, 7)
      .setBorder(true, true, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  // Fila ACTIVIDAD — ocupa todo el ancho restante con wrap para textos largos
  sheet.setRowHeight(r, 36);
  sheet.getRange(r, 1, 1, 2).merge().setValue("ACTIVIDAD")
    .setFontSize(8).setFontWeight("bold").setFontColor("#555555")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBackground("#f5f5f5");
  sheet.getRange(r, 3, 1, 5).merge().setValue(apu.actividad || "—")
    .setFontSize(9).setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setWrap(true).setBackground("#ffffff");
  sheet.getRange(r, 1, 1, 7)
    .setBorder(true, true, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
  r++;

  // Secciones del APU
  const secciones = [
    { key: "equipos",    label: "A — EQUIPOS",      campo: "subtotal_equipos"    },
    { key: "materiales", label: "B — MATERIALES",   campo: "subtotal_materiales" },
    { key: "mano_obra",  label: "C — MANO DE OBRA", campo: "subtotal_mano_obra"  },
    { key: "otros",      label: "D — OTROS",        campo: "subtotal_otros"      },
  ];

  secciones.forEach(sec => {
    const items    = apu[sec.key] || [];
    const subtotal = parseFloat(apu[sec.campo]) || 0;

    // Encabezado de sección — con sufijo informativo según tipo
    let labelSec = sec.label;
    if (sec.key === "materiales" && desperdicioPct > 0) {
      labelSec += "   (Incluye factor de desperdicio del " + (desperdicioPct % 1 === 0 ? desperdicioPct.toFixed(0) : desperdicioPct.toString()) + "%)";
    } else if (sec.key === "mano_obra") {
      labelSec += "   " + construirSufijoPrestaciones(apu.mano_obra, manoObraPrestById);
    }
    sheet.setRowHeight(r, 24);
    sheet.getRange(r, 1, 1, 7).merge()
      .setValue(labelSec)
      .setBackground(GRIS_OSC).setFontColor("#333333")
      .setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("left").setVerticalAlignment("middle");
    r++;

    const renderHM = sec.key === "equipos" && hmPct > 0;
    if (items.length > 0 || renderHM) {
      // Cabecera de columnas
      const showUnidad = sec.key === "materiales";
      sheet.setRowHeight(r, 20);
      if (showUnidad) {
        sheet.getRange(r, 1, 1, 7)
          .setValues([["ÍTEM", "DESCRIPCIÓN", "UNIDAD", "CANT.", "REND.", "P. UNITARIO", "VALOR PARCIAL"]])
          .setBackground(GRIS_CLR).setFontColor("#333333")
          .setFontWeight("bold").setFontSize(8)
          .setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBorder(true, true, true, true, true, true, BORDE, SpreadsheetApp.BorderStyle.SOLID);
      } else {
        sheet.getRange(r, 1).setValue("ÍTEM")
          .setBackground(GRIS_CLR).setFontColor("#333333").setFontWeight("bold").setFontSize(8)
          .setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBorder(true, true, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 2, 1, 2).merge().setValue("DESCRIPCIÓN")
          .setBackground(GRIS_CLR).setFontColor("#333333").setFontWeight("bold").setFontSize(8)
          .setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        ["CANT.", "REND.", "P. UNITARIO", "VALOR PARCIAL"].forEach((h, i) => {
          sheet.getRange(r, 4 + i).setValue(h)
            .setBackground(GRIS_CLR).setFontColor("#333333").setFontWeight("bold").setFontSize(8)
            .setHorizontalAlignment("center").setVerticalAlignment("middle")
            .setBorder(true, null, true, i === 3, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        });
      }
      r++;

      items.forEach((it, idx) => {
        const rend = (it.rendimiento !== null && it.rendimiento !== "") ? it.rendimiento : "—";
        const bg   = idx % 2 === 0 ? "#ffffff" : "#fafafa";
        const codigo = resolverCodigoItem(it, materialesCodigoById);
        sheet.getRange(r, 1).setValue(codigo)
          .setFontSize(8).setHorizontalAlignment("center").setVerticalAlignment("middle");
        if (showUnidad) {
          sheet.getRange(r, 2).setValue(it.descripcion_manual || "—")
            .setFontSize(8).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true);
          sheet.getRange(r, 3).setValue(it.unidad || "")
            .setFontSize(8).setHorizontalAlignment("center").setVerticalAlignment("middle");
        } else {
          sheet.getRange(r, 2, 1, 2).merge().setValue(it.descripcion_manual || "—")
            .setFontSize(8).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true);
        }
        sheet.getRange(r, 4).setValue(it.cantidad || 0)
          .setFontSize(8).setHorizontalAlignment("right").setVerticalAlignment("middle");
        sheet.getRange(r, 5).setValue(rend)
          .setFontSize(8).setHorizontalAlignment("right").setVerticalAlignment("middle");
        sheet.getRange(r, 6).setValue(parseFloat(it.precio_unitario) || 0)
          .setNumberFormat(MONEY).setFontSize(8).setHorizontalAlignment("right").setVerticalAlignment("middle");
        const vpBase     = parseFloat(it.valor_parcial) || 0;
        const vpMostrado = sec.key === "materiales" ? vpBase * desperdicioFactor : vpBase;
        sheet.getRange(r, 7).setValue(vpMostrado)
          .setNumberFormat(MONEY).setFontSize(8).setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
        sheet.getRange(r, 1, 1, 7)
          .setBackground(bg)
          .setBorder(true, true, true, true, true, true, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        r++;
      });

      // Línea automática Herramienta Menor (HM) al final de Equipos
      if (renderHM) {
        const bg = items.length % 2 === 0 ? "#ffffff" : "#fafafa";
        sheet.setRowHeight(r, 20);
        sheet.getRange(r, 1).setValue("HM")
          .setFontSize(8).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
        sheet.getRange(r, 2, 1, 2).merge().setValue("Herramienta Menor (% MO)")
          .setFontSize(8).setHorizontalAlignment("left").setVerticalAlignment("middle");
        sheet.getRange(r, 4).setValue(hmPct / 100)
          .setNumberFormat("0.##%").setFontSize(8).setHorizontalAlignment("right").setVerticalAlignment("middle");
        sheet.getRange(r, 5).setValue("—")
          .setFontSize(8).setHorizontalAlignment("right").setVerticalAlignment("middle");
        sheet.getRange(r, 6).setValue(subtotalMOBase)
          .setNumberFormat(MONEY).setFontSize(8).setHorizontalAlignment("right").setVerticalAlignment("middle");
        sheet.getRange(r, 7).setValue(hmValor)
          .setNumberFormat(MONEY).setFontSize(8).setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
        sheet.getRange(r, 1, 1, 7)
          .setBackground(bg)
          .setBorder(true, true, true, true, true, true, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        r++;
      }
    } else {
      sheet.setRowHeight(r, 18);
      sheet.getRange(r, 1, 1, 7).merge()
        .setValue("Sin ítems").setFontSize(8).setFontColor("#aaaaaa")
        .setHorizontalAlignment("center").setBackground("#fafafa");
      r++;
    }

    // Subtotal de sección
    sheet.setRowHeight(r, 22);
    sheet.getRange(r, 1, 1, 6).merge()
      .setValue("SUBTOTAL " + sec.label.split("—")[0].trim())
      .setFontSize(9).setFontWeight("bold").setFontColor("#333333")
      .setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBackground(GRIS_CLR)
      .setBorder(true, true, true, null, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 7).setValue(subtotal)
      .setNumberFormat(MONEY).setFontSize(9).setFontWeight("bold").setFontColor("#111111")
      .setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBackground(GRIS_CLR)
      .setBorder(true, null, true, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  // Separador antes del total
  sheet.setRowHeight(r, 4);
  sheet.getRange(r, 1, 1, 7).merge().setValue("").setBackground("#888888");
  r++;

  // COSTO DIRECTO TOTAL
  sheet.setRowHeight(r, 24);
  const costoNeto = parseFloat(apu.costo_neto) || 0;
  sheet.getRange(r, 1, 1, 6).merge()
    .setValue("COSTOS DIRECTOS  (A + B + C + D)")
    .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle")
    .setBackground(NEGRO);
  sheet.getRange(r, 7).setValue(costoNeto)
    .setNumberFormat(MONEY).setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle")
    .setBackground(NEGRO);
  r++;

  // Encabezado COSTOS INDIRECTOS
  sheet.setRowHeight(r, 18);
  sheet.getRange(r, 1, 1, 7).merge()
    .setValue("COSTOS INDIRECTOS")
    .setFontSize(9).setFontWeight("bold").setFontColor("#333333")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBackground("#c8c8c8");
  r++;

  // Filas de indirectos
  const admPct  = parseFloat(apu.administracion_pct) || 0;
  const impPct  = parseFloat(apu.imprevistos_pct)    || 0;
  const utilPct = parseFloat(apu.utilidad_pct)       || 0;
  const indirectos = [
    ["Administración",  admPct,  Math.round(costoNeto * admPct  / 100)],
    ["Imprevistos",     impPct,  Math.round(costoNeto * impPct  / 100)],
    ["Utilidad",        utilPct, Math.round(costoNeto * utilPct / 100)],
  ];
  indirectos.forEach(([label, pct, valor]) => {
    sheet.setRowHeight(r, 20);
    sheet.getRange(r, 1, 1, 4).merge()
      .setValue(label)
      .setFontSize(9).setFontColor("#333333")
      .setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBackground(GRIS_CLR)
      .setBorder(true, true, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 5, 1, 2).merge()
      .setValue(pct / 100)
      .setNumberFormat("0.##%").setFontSize(9).setFontColor("#555555")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBackground(GRIS_CLR)
      .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 7).setValue(valor)
      .setNumberFormat(MONEY).setFontSize(9).setFontColor("#333333")
      .setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBackground(GRIS_CLR)
      .setBorder(true, null, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  // VALOR UNITARIO TOTAL
  const subtotalInd = Math.round(costoNeto * (admPct + impPct + utilPct) / 100);
  const valorTotal  = costoNeto + subtotalInd;
  sheet.setRowHeight(r, 30);
  sheet.getRange(r, 1, 1, 6).merge()
    .setValue("VALOR UNITARIO TOTAL")
    .setFontSize(12).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle")
    .setBackground(NEGRO);
  sheet.getRange(r, 7).setValue(valorTotal)
    .setNumberFormat(MONEY).setFontSize(12).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle")
    .setBackground(NEGRO);
}

// ─── ELIMINAR APU ────────────────────────────────────────────────────────────

function eliminarAPU(apuId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const itemsSheet = ss.getSheetByName("APU_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();
  const ih         = itemsData[0];
  for (let i = itemsData.length - 1; i >= 1; i--) {
    if (itemsData[i][ih.indexOf("apu_id")] == apuId) itemsSheet.deleteRow(i + 1);
  }

  const apuSheet = ss.getSheetByName("APU");
  const apuData  = apuSheet.getDataRange().getValues();
  for (let i = 1; i < apuData.length; i++) {
    if (apuData[i][0] == apuId) { apuSheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

// ─── LISTA DE APUs ────────────────────────────────────────────────────────────

function listarAPUs() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU");

  // Si la hoja APU está vacía pero hay datos en otras tablas, reconstruir automáticamente
  if (!sheet || sheet.getLastRow() < 2) {
    reconstruirAPUsDesdeItems();
  }

  return sheetToObjects(ss, "APU").map(a => ({
    id:          a.id,
    codigo_item: a.codigo_item,
    descripcion: a.descripcion,
    unidad:      a.unidad,
    cliente:     a.cliente,
    direccion:   a.direccion || "",
    actividad:   a.actividad,
    costo_neto:  a.costo_neto,
    fecha:       String(a.fecha || ""),
  }));
}

// ─── RECONSTRUIR APUs DESDE DATOS EXISTENTES ──────────────────────────────────
// Recupera APUs usando Cotizacion_Items (descripcion, cliente, costo) + APU_Items (subtotales)

function reconstruirAPUsDesdeItems() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const apuSheet = ss.getSheetByName("APU");

  if (!apuSheet) return { ok: false, error: "Hoja 'APU' no encontrada." };
  if (apuSheet.getLastRow() >= 2) return { ok: false, error: "La hoja APU ya tiene datos." };

  // Recolectar IDs desde APU_Items y Cotizacion_Items
  const apuItemsAll = sheetToObjects(ss, "APU_Items");
  const cotItemsAll = sheetToObjects(ss, "Cotizacion_Items");

  const apuIds = new Set();
  apuItemsAll.forEach(i => { if (i.apu_id) apuIds.add(String(i.apu_id)); });
  cotItemsAll.forEach(i => { if (i.apu_id) apuIds.add(String(i.apu_id)); });

  if (!apuIds.size) return { ok: false, error: "No se encontraron APUs en APU_Items ni en Cotizacion_Items. No hay datos para recuperar." };

  // Lookup de cotizaciones para obtener el cliente
  const cotizaciones = sheetToObjects(ss, "Cotizaciones");
  const cotById = {};
  cotizaciones.forEach(c => { cotById[String(c.id)] = c; });

  // Mejor info de cada APU desde Cotizacion_Items (descripcion, unidad, cliente, precio)
  const infoByApu = {};
  cotItemsAll.forEach(ci => {
    const aid = String(ci.apu_id);
    if (!infoByApu[aid]) {
      const cot = cotById[String(ci.cotizacion_id)] || {};
      infoByApu[aid] = {
        descripcion: ci.descripcion || "",
        unidad:      ci.unidad      || "",
        precio_apu:  parseFloat(ci.precio_apu) || 0,
        cliente:     cot.cliente    || "",
      };
    }
  });

  // Subtotales desde APU_Items
  const subsByApu = {};
  apuItemsAll.forEach(item => {
    const aid = String(item.apu_id);
    if (!subsByApu[aid]) subsByApu[aid] = { EQUIPO: 0, MATERIAL: 0, MANO_OBRA: 0, OTRO: 0 };
    subsByApu[aid][item.tipo] = (subsByApu[aid][item.tipo] || 0) + (parseFloat(item.valor_parcial) || 0);
  });

  const apuHeaders = apuSheet.getRange(1, 1, 1, apuSheet.getLastColumn()).getValues()[0];
  const fecha      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  [...apuIds].sort((a, b) => Number(a) - Number(b)).forEach(apuId => {
    const info   = infoByApu[apuId] || {};
    const subs   = subsByApu[apuId] || {};
    const subEq  = subs.EQUIPO    || 0;
    const subMat = subs.MATERIAL  || 0;
    const subMo  = subs.MANO_OBRA || 0;
    const subOt  = subs.OTRO      || 0;
    const neto   = info.precio_apu || (subEq + subMat + subMo + subOt);

    const valores = {
      id:                    Number(apuId),
      codigo_item:           info.descripcion || ("APU-" + apuId),
      descripcion:           info.descripcion || ("APU-" + apuId),
      unidad:                info.unidad      || "",
      cliente:               info.cliente     || "",
      direccion:             "",
      actividad:             info.descripcion || "",
      subtotal_equipos:      subEq,
      subtotal_materiales:   subMat,
      subtotal_mano_obra:    subMo,
      subtotal_otros:        subOt,
      costo_neto:            neto,
      administracion_pct:    0,
      imprevistos_pct:       0,
      utilidad_pct:          0,
      iva_pct:               19,
      valor_total:           neto,
      fecha:                 fecha,
      desperdicio_pct:       0,
      herramienta_menor_pct: 0,
    };
    apuSheet.appendRow(apuHeaders.map(h => (valores[h] !== undefined ? valores[h] : "")));
  });

  return { ok: true, count: apuIds.size };
}

// ─── CREAR APU ────────────────────────────────────────────────────────────────

function crearAPU(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU");
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId = lastId + 1;
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  const valores = {
    id:                    newId,
    codigo_item:           datos.codigo_item || "",
    descripcion:           datos.descripcion || datos.codigo_item || "",
    unidad:                datos.unidad      || "",
    cliente:               datos.cliente     || "",
    direccion:             datos.direccion   || "",
    actividad:             datos.actividad   || "",
    subtotal_equipos:      0,
    subtotal_materiales:   0,
    subtotal_mano_obra:    0,
    subtotal_otros:        0,
    costo_neto:            0,
    administracion_pct:    0,
    imprevistos_pct:       0,
    utilidad_pct:          0,
    iva_pct:               19,
    valor_total:           0,
    fecha:                 fecha,
    desperdicio_pct:       0,
    herramienta_menor_pct: 0,
  };
  sheet.appendRow(headers.map(h => valores[h] !== undefined ? valores[h] : ""));
  return newId;
}

// ─── ACTUALIZAR CABEZA DEL APU ────────────────────────────────────────────────

function actualizarCabezaAPU(apuId, datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU");
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == apuId) {
      const campos = [
        "codigo_item", "descripcion", "unidad", "cliente", "direccion", "actividad",
        "administracion_pct", "imprevistos_pct", "utilidad_pct", "iva_pct",
        "desperdicio_pct", "herramienta_menor_pct",
      ];
      campos.forEach(campo => {
        if (datos[campo] !== undefined) {
          const col = headers.indexOf(campo);
          if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(datos[campo]);
        }
      });
      // Si cambió desperdicio_pct o herramienta_menor_pct, recalcular subtotales (afecta neto + valor_total)
      if (datos.desperdicio_pct !== undefined || datos.herramienta_menor_pct !== undefined) {
        recalcularSubtotales(ss, apuId);
      } else if (datos.administracion_pct !== undefined || datos.imprevistos_pct !== undefined ||
                 datos.utilidad_pct !== undefined       || datos.iva_pct        !== undefined) {
        // Solo cambió AIU/IVA — recomputar valor_total sin tocar subtotales
        const row = data[i];
        const neto      = parseFloat(row[headers.indexOf("costo_neto")])        || 0;
        const admPct    = parseFloat(datos.administracion_pct ?? row[headers.indexOf("administracion_pct")]) || 0;
        const impPct    = parseFloat(datos.imprevistos_pct    ?? row[headers.indexOf("imprevistos_pct")])    || 0;
        const utilPct   = parseFloat(datos.utilidad_pct       ?? row[headers.indexOf("utilidad_pct")])       || 0;
        const ivaPct    = parseFloat(datos.iva_pct            ?? row[headers.indexOf("iva_pct")])            || 0;
        const subtotalAIU = neto * (admPct + impPct + utilPct) / 100;
        const sinIVA      = neto + subtotalAIU;
        const iva         = sinIVA * ivaPct / 100;
        const total       = sinIVA + iva;
        const colVT = headers.indexOf("valor_total");
        if (colVT >= 0) sheet.getRange(i + 1, colVT + 1).setValue(Math.round(total));
      }
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─── GET APU COMPLETO ─────────────────────────────────────────────────────────

function getAPUCompleto(apuId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const apuSheet   = ss.getSheetByName("APU");
  const apuData    = apuSheet.getDataRange().getValues();
  const apuHeaders = apuData[0];
  const apuRow     = apuData.slice(1).find(r => r[0] == apuId);
  if (!apuRow) return null;

  const apu = {};
  apuHeaders.forEach((h, i) => {
    const v = apuRow[i];
    apu[h] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy") : v;
  });

  const itemsSheet = ss.getSheetByName("APU_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();
  if (itemsData.length < 2) {
    apu.equipos = []; apu.materiales = []; apu.mano_obra = []; apu.otros = [];
    return apu;
  }

  const ih    = itemsData[0];
  const items = itemsData.slice(1)
    .filter(r => r[ih.indexOf("apu_id")] == apuId)
    .map(r => {
      const o = {};
      ih.forEach((h, i) => {
        const v = r[i];
        o[h] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy") : v;
      });
      return o;
    });

  // Enrich items with unidad from BD (APU_Items doesn't store it)
  const matUnidad  = {};
  sheetToObjects(ss, "Materiales").forEach(m => { matUnidad[String(m.id)] = m.unidad || ""; });
  const otroUnidad = {};
  sheetToObjects(ss, "Otros").forEach(o => { otroUnidad[String(o.id)] = o.unidad || ""; });
  items.forEach(item => {
    if (!item.unidad) {
      if      (item.tipo === "MATERIAL")                              item.unidad = matUnidad[String(item.recurso_id)] || "";
      else if (item.tipo === "EQUIPO" || item.tipo === "MANO_OBRA")  item.unidad = "Día";
      else if (item.tipo === "OTRO")                                  item.unidad = otroUnidad[String(item.recurso_id)] || "";
    }
  });

  apu.equipos    = items.filter(i => i.tipo === "EQUIPO");
  apu.materiales = items.filter(i => i.tipo === "MATERIAL");
  apu.mano_obra  = items.filter(i => i.tipo === "MANO_OBRA");
  apu.otros      = items.filter(i => i.tipo === "OTRO");
  return apu;
}

// ─── AGREGAR ÍTEM AL APU ──────────────────────────────────────────────────────

function agregarItemAPU(apuId, item) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU_Items");
  const data  = sheet.getDataRange().getValues();
  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId = lastId + 1;

  const cant    = parseFloat(item.cantidad)        || 0;
  const precio  = parseFloat(item.precio_unitario) || 0;
  const rend    = parseFloat(item.rendimiento)     || 1;
  const vp      = calcularValorParcial(item.tipo, cant, precio, rend);
  const partida = getPartidaRecurso(ss, item.tipo, item.recurso_id || "");

  sheet.appendRow([
    newId, apuId, item.tipo,
    item.recurso_id         || "",
    item.descripcion_manual || "",
    cant, rend, precio, vp, partida
  ]);

  recalcularSubtotales(ss, apuId);
  return { id: newId, valor_parcial: vp };
}

// ─── ELIMINAR ÍTEM ────────────────────────────────────────────────────────────

function eliminarItemAPU(itemId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU_Items");
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const apuId = data[i][headers.indexOf("apu_id")];
      sheet.deleteRow(i + 1);
      recalcularSubtotales(ss, apuId);
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─── ACTUALIZAR ÍTEM ──────────────────────────────────────────────────────────

function actualizarItemAPU(itemId, cambios) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU_Items");
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const row   = [...data[i]];
      const tipo  = row[headers.indexOf("tipo")];
      const apuId = row[headers.indexOf("apu_id")];

      const cant  = parseFloat(cambios.cantidad        ?? row[headers.indexOf("cantidad")])        || 0;
      const precio= parseFloat(cambios.precio_unitario ?? row[headers.indexOf("precio_unitario")]) || 0;
      const rend  = parseFloat(cambios.rendimiento     ?? row[headers.indexOf("rendimiento")])     || 1;
      const vp    = calcularValorParcial(tipo, cant, precio, rend);

      if (cambios.cantidad         !== undefined) row[headers.indexOf("cantidad")]          = cant;
      if (cambios.precio_unitario  !== undefined) row[headers.indexOf("precio_unitario")]   = precio;
      if (cambios.rendimiento      !== undefined) row[headers.indexOf("rendimiento")]       = rend;
      if (cambios.descripcion_manual !== undefined) row[headers.indexOf("descripcion_manual")] = cambios.descripcion_manual;
      row[headers.indexOf("valor_parcial")] = vp;

      sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      recalcularSubtotales(ss, apuId);
      return { ok: true, valor_parcial: vp };
    }
  }
  return { ok: false };
}

// ─── ACTUALIZAR PRECIO EN BD ──────────────────────────────────────────────────

function actualizarPrecioRecurso(tipo, id, nuevoPrecio) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const mapa = { MATERIAL: "Materiales", EQUIPO: "Equipos", MANO_OBRA: "ManoObra" };
  if (!mapa[tipo]) return { ok: false };

  const sheet   = ss.getSheetByName(mapa[tipo]);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  const colNombre = tipo === "MANO_OBRA" ? "salario_mensual"
                  : tipo === "EQUIPO"    ? "tarifa_dia"
                  :                        "precio_sin_iva";
  const col = headers.indexOf(colNombre);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, col + 1).setValue(nuevoPrecio);
      if (tipo === "MANO_OBRA") {
        const pct     = data[i][headers.indexOf("prestaciones_pct")] || 54;
        const costoDia= Math.round(nuevoPrecio * (1 + pct / 100) / 30);
        sheet.getRange(i + 1, headers.indexOf("costo_dia") + 1).setValue(costoDia);
      }
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─── ACTUALIZAR PRESTACIONES SOCIALES A TODOS LOS ROLES ──────────────────────
// Aplica el mismo % a todos los roles de ManoObra y recalcula costo_dia.
// Útil cuando cambia la ley colombiana de prestaciones (54% → otro valor).
// No toca precio_unitario de items ya guardados en APU_Items.

function actualizarPrestacionesPctMasivo(nuevoPct) {
  const pct = parseFloat(nuevoPct);
  if (isNaN(pct) || pct < 0 || pct > 200) return { ok: false, error: "Porcentaje inválido (debe estar entre 0 y 200)." };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ManoObra");
  if (!sheet) return { ok: false, error: "Hoja ManoObra no encontrada." };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxSal    = headers.indexOf("salario_mensual");
  const idxPrest  = headers.indexOf("prestaciones_pct");
  const idxCosto  = headers.indexOf("costo_dia");
  if (idxSal < 0 || idxPrest < 0 || idxCosto < 0) {
    return { ok: false, error: "Faltan columnas requeridas en ManoObra." };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const factor = 1 + pct / 100;
  data.forEach(row => {
    const sal = parseFloat(row[idxSal]) || 0;
    row[idxPrest] = pct;
    row[idxCosto] = Math.round(sal * factor / 30);
  });
  sheet.getRange(2, 1, data.length, sheet.getLastColumn()).setValues(data);

  return { ok: true, count: data.length };
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

// Fórmulas exactas del APU del ingeniero:
// Equipos:   (a × b) / c  = (cant × tarifa_dia) / rendimiento
// Materiales: d × e        = precio × cantidad
// Mano obra: (f × g) / h  = (cant × costo_dia) / rendimiento
// Otros:      i × j        = costo_unitario × cantidad
function calcularValorParcial(tipo, cant, precio, rend) {
  if (tipo === "EQUIPO"   || tipo === "MANO_OBRA") return rend > 0 ? (cant * precio) / rend : 0;
  if (tipo === "MATERIAL")                          return cant * precio;
  if (tipo === "OTRO")                              return precio * cant;
  return 0;
}

function _padId(id) {
  const n = parseInt(id);
  if (isNaN(n)) return String(id);
  return String(n).padStart(3, "0");
}

// Devuelve el código a mostrar en la columna ÍTEM del APU.
// MATERIAL → Materiales.codigo (lookup en mapa precargado). Si falta, fallback a MAT-<id>.
// EQUIPO/MANO_OBRA/OTRO → autogenerado con prefijo + id zero-padded a 3.
// Ítem sin recurso_id (descripción manual) → "—".
// Construye el sufijo del encabezado de Mano de Obra mostrando el % real de prestaciones
// según los roles utilizados en los ítems. Si los roles usan el mismo %, lo muestra;
// si hay varios, dice "según rol".
function construirSufijoPrestaciones(moItems, manoObraPrestById) {
  const set = new Set();
  (moItems || []).forEach(it => {
    const pct = manoObraPrestById ? manoObraPrestById[String(it.recurso_id)] : undefined;
    if (pct !== undefined && pct !== null && pct !== "" && !isNaN(pct)) set.add(parseFloat(pct));
  });
  if (set.size === 0) return "(Prestaciones Sociales)";
  if (set.size === 1) {
    const pct = [...set][0];
    return "(Prestaciones Sociales del " + (pct % 1 === 0 ? pct.toFixed(0) : pct.toString()) + "%)";
  }
  return "(Prestaciones Sociales según rol)";
}

function resolverCodigoItem(item, materialesCodigoById) {
  if (!item) return "—";
  const id = item.recurso_id;
  if (id === "" || id === null || id === undefined) return "—";
  switch (item.tipo) {
    case "MATERIAL": {
      const codigo = materialesCodigoById ? materialesCodigoById[String(id)] : "";
      return codigo || ("MAT-" + _padId(id));
    }
    case "EQUIPO":    return "EQ-" + _padId(id);
    case "MANO_OBRA": return "MO-" + _padId(id);
    case "OTRO":      return "OT-" + _padId(id);
    default:          return "—";
  }
}

function recalcularSubtotales(ss, apuId) {
  const itemsSheet = ss.getSheetByName("APU_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();

  const apuSheet = ss.getSheetByName("APU");
  const apuData  = apuSheet.getDataRange().getValues();
  const ah       = apuData[0];

  // Localizar fila del APU
  let i = -1;
  for (let k = 1; k < apuData.length; k++) {
    if (apuData[k][0] == apuId) { i = k; break; }
  }
  if (i < 0) return;

  // Sumar valores parciales base por tipo
  let baseEq = 0, baseMat = 0, subMo = 0, subOt = 0;
  if (itemsData.length >= 2) {
    const h = itemsData[0];
    const idxApu  = h.indexOf("apu_id");
    const idxTipo = h.indexOf("tipo");
    const idxVP   = h.indexOf("valor_parcial");
    for (let k = 1; k < itemsData.length; k++) {
      const r = itemsData[k];
      if (r[idxApu] != apuId) continue;
      const vp = parseFloat(r[idxVP]) || 0;
      switch (r[idxTipo]) {
        case "EQUIPO":    baseEq  += vp; break;
        case "MATERIAL":  baseMat += vp; break;
        case "MANO_OBRA": subMo   += vp; break;
        case "OTRO":      subOt   += vp; break;
      }
    }
  }

  // Leer % desperdicio y HM (columnas pueden no existir en hojas pre-migración)
  const idxDesp = ah.indexOf("desperdicio_pct");
  const idxHm   = ah.indexOf("herramienta_menor_pct");
  const desperdicioPct = idxDesp >= 0 ? (parseFloat(apuData[i][idxDesp]) || 0) : 0;
  const hmPct          = idxHm   >= 0 ? (parseFloat(apuData[i][idxHm])   || 0) : 0;

  const subMat  = baseMat * (1 + desperdicioPct / 100);
  const hmValor = subMo * hmPct / 100;
  const subEq   = baseEq + hmValor;
  const neto    = subEq + subMat + subMo + subOt;

  apuSheet.getRange(i + 1, ah.indexOf("subtotal_equipos")    + 1).setValue(subEq);
  apuSheet.getRange(i + 1, ah.indexOf("subtotal_materiales") + 1).setValue(subMat);
  apuSheet.getRange(i + 1, ah.indexOf("subtotal_mano_obra")  + 1).setValue(subMo);
  apuSheet.getRange(i + 1, ah.indexOf("subtotal_otros")      + 1).setValue(subOt);
  apuSheet.getRange(i + 1, ah.indexOf("costo_neto")          + 1).setValue(neto);

  const admPct  = parseFloat(apuData[i][ah.indexOf("administracion_pct")]) || 0;
  const impPct  = parseFloat(apuData[i][ah.indexOf("imprevistos_pct")])    || 0;
  const utilPct = parseFloat(apuData[i][ah.indexOf("utilidad_pct")])       || 0;
  const ivaPct  = parseFloat(apuData[i][ah.indexOf("iva_pct")])            || 0;
  const subtAIU = neto * (admPct + impPct + utilPct) / 100;
  const sinIVA  = neto + subtAIU;
  const total   = sinIVA + sinIVA * ivaPct / 100;
  const colVT = ah.indexOf("valor_total");
  if (colVT >= 0) apuSheet.getRange(i + 1, colVT + 1).setValue(Math.round(total));
}

// ─── MIGRACIÓN AIU EN APU ─────────────────────────────────────────────────────
// Agrega columnas AIU a la hoja APU si no existen todavía.

function migrarAIUenAPU() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU");
  if (!sheet) { SpreadsheetApp.getUi().alert("Hoja APU no encontrada."); return; }

  const lastCol = sheet.getLastColumn();
  const hRow    = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const nuevas  = ["administracion_pct", "imprevistos_pct", "utilidad_pct", "iva_pct", "valor_total"];
  const msgs    = [];

  nuevas.forEach(col => {
    if (!hRow.includes(col)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col);
      if (col === "iva_pct" && sheet.getLastRow() > 1) {
        sheet.getRange(2, newCol, sheet.getLastRow() - 1, 1).setValue(19);
      }
      msgs.push("✅ Columna '" + col + "' agregada.");
    } else {
      msgs.push("— '" + col + "' ya existía.");
    }
  });

  SpreadsheetApp.getUi().alert("Migración AIU en APU:\n\n" + msgs.join("\n"));
}

function sheetToObjects(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
    return obj;
  });
}
