// ─── CARGA DE BD (para búsqueda client-side) ─────────────────────────────────

function cargarTodosBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    materiales: sheetToObjects(ss, "Materiales"),
    manoObra:   sheetToObjects(ss, "ManoObra"),
    equipos:    sheetToObjects(ss, "Equipos"),
  };
}

// ─── LISTA DE APUs ────────────────────────────────────────────────────────────

function listarAPUs() {
  return sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "APU");
}

// ─── CREAR APU ────────────────────────────────────────────────────────────────

function crearAPU(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("APU");
  const data  = sheet.getDataRange().getValues();
  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId = lastId + 1;
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  sheet.appendRow([
    newId,
    datos.codigo_item || "",
    datos.descripcion || datos.codigo_item || "",
    datos.unidad      || "",
    datos.cliente     || "",
    datos.actividad   || "",
    0, 0, 0, 0, 0,
    fecha
  ]);
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
      const campos = ["codigo_item", "descripcion", "unidad", "cliente", "actividad"];
      campos.forEach(campo => {
        if (datos[campo] !== undefined) {
          const col = headers.indexOf(campo);
          if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(datos[campo]);
        }
      });
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
  apuHeaders.forEach((h, i) => apu[h] = apuRow[i]);

  const itemsSheet   = ss.getSheetByName("APU_Items");
  const itemsData    = itemsSheet.getDataRange().getValues();
  if (itemsData.length < 2) {
    apu.equipos = []; apu.materiales = []; apu.mano_obra = []; apu.otros = [];
    return apu;
  }

  const ih    = itemsData[0];
  const items = itemsData.slice(1)
    .filter(r => r[ih.indexOf("apu_id")] == apuId)
    .map(r => { const o = {}; ih.forEach((h, i) => o[h] = r[i]); return o; });

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

  const cant  = parseFloat(item.cantidad)        || 0;
  const precio= parseFloat(item.precio_unitario) || 0;
  const rend  = parseFloat(item.rendimiento)     || 1;
  const vp    = calcularValorParcial(item.tipo, cant, precio, rend);

  sheet.appendRow([
    newId, apuId, item.tipo,
    item.recurso_id        || "",
    item.descripcion_manual|| "",
    cant, rend, precio, vp
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
// Otros:      i / j        = valor_total / cantidad_unidades
function calcularValorParcial(tipo, cant, precio, rend) {
  if (tipo === "EQUIPO"   || tipo === "MANO_OBRA") return rend > 0 ? (cant * precio) / rend : 0;
  if (tipo === "MATERIAL")                          return cant * precio;
  if (tipo === "OTRO")                              return cant > 0 ? precio / cant : 0;
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
      break;
    }
  }
}

function sheetToObjects(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}
