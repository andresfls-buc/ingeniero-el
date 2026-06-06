// ─── COTIZACIÓN CLIENTE — CRUD ────────────────────────────────────────────────
// Hoja: Cotizacion_Cliente_Items
// Columnas: id, cotizacion_id, item_num, descripcion, unidad, cantidad, precio_unitario, valor_total

function _clienteSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cotizacion_Cliente_Items");
  if (!sheet) throw new Error("Hoja Cotizacion_Cliente_Items no encontrada. Ejecuta migrarHojaClienteItems() primero.");
  return sheet;
}

// Devuelve { cot: {...}, items: [...], cfg: {...} }
function getCotizacionClienteDoc(cotId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Cabecera de la cotización (cliente, fecha, numero_oferta, AIU, condiciones)
  const cotSheet = ss.getSheetByName("Cotizaciones");
  const cotData  = cotSheet.getDataRange().getValues();
  const cotH     = cotData[0];
  const cotRow   = cotData.slice(1).find(r => r[0] == cotId);
  if (!cotRow) return null;
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

  return { cot, items, cfg };
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
