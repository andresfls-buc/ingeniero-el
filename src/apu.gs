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
  const items = sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "APU");
  return items.map(a => ({
    id:          a.id,
    codigo_item: a.codigo_item,
    descripcion: a.descripcion,
    unidad:      a.unidad,
    cliente:     a.cliente,
    direccion:   a.direccion   || "",
    actividad:   a.actividad,
    costo_neto:  a.costo_neto,
    fecha:       String(a.fecha || ""),
  }));
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
    id:                  newId,
    codigo_item:         datos.codigo_item || "",
    descripcion:         datos.descripcion || datos.codigo_item || "",
    unidad:              datos.unidad      || "",
    cliente:             datos.cliente     || "",
    direccion:           datos.direccion   || "",
    actividad:           datos.actividad   || "",
    subtotal_equipos:    0,
    subtotal_materiales: 0,
    subtotal_mano_obra:  0,
    subtotal_otros:      0,
    costo_neto:          0,
    administracion_pct:  0,
    imprevistos_pct:     0,
    utilidad_pct:        0,
    iva_pct:             19,
    valor_total:         0,
    fecha:               fecha,
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
      ];
      campos.forEach(campo => {
        if (datos[campo] !== undefined) {
          const col = headers.indexOf(campo);
          if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(datos[campo]);
        }
      });
      // Recalculate valor_total if any AIU field was updated
      if (datos.administracion_pct !== undefined || datos.imprevistos_pct !== undefined ||
          datos.utilidad_pct !== undefined       || datos.iva_pct        !== undefined) {
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

function recalcularSubtotales(ss, apuId) {
  const itemsSheet = ss.getSheetByName("APU_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();
  if (itemsData.length < 2) return;

  const h     = itemsData[0];
  const items = itemsData.slice(1).filter(r => r[h.indexOf("apu_id")] == apuId);
  const sum   = tipo => items
    .filter(r => r[h.indexOf("tipo")] === tipo)
    .reduce((acc, r) => acc + (parseFloat(r[h.indexOf("valor_parcial")]) || 0), 0);

  const subEq  = sum("EQUIPO");
  const subMat = sum("MATERIAL");
  const subMo  = sum("MANO_OBRA");
  const subOt  = sum("OTRO");
  const neto   = subEq + subMat + subMo + subOt;

  const apuSheet = ss.getSheetByName("APU");
  const apuData  = apuSheet.getDataRange().getValues();
  const ah       = apuData[0];

  for (let i = 1; i < apuData.length; i++) {
    if (apuData[i][0] == apuId) {
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
      break;
    }
  }
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
