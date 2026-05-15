// ─── CRUD BASE DE DATOS DE PRECIOS ───────────────────────────────────────────

// Carga todos los campos de Materiales (incluye proveedor, para el panel de gestión).
function cargarMaterialesBD() {
  return sheetToObjects(SpreadsheetApp.getActiveSpreadsheet(), "Materiales").map(m => ({
    id:                  m.id,
    codigo:              m.codigo              || "",
    categoria:           m.categoria           || "",
    nombre:              m.nombre              || "",
    unidad:              m.unidad              || "",
    precio_sin_iva:      parseFloat(m.precio_sin_iva) || 0,
    proveedor:           m.proveedor           || "",
    fecha_actualizacion: m.fecha_actualizacion || "",
  }));
}

// Crea un nuevo registro en la hoja correspondiente al tipo.
function crearRecurso(tipo, datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const conf  = _bdConf(tipo);
  const sheet = ss.getSheetByName(conf.hoja);
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];
  const lastId = data.length > 1 ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0)) : 0;
  const newId  = lastId + 1;
  const row    = h.map(col => conf.campoValor(col, newId, datos));
  sheet.appendRow(row);
  return { id: newId, ok: true };
}

// Actualiza todos los campos editables de un registro existente.
function actualizarRecurso(tipo, id, datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const conf  = _bdConf(tipo);
  const sheet = ss.getSheetByName(conf.hoja);
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const row = [...data[i]];
      h.forEach((col, ci) => {
        const val = conf.campoValorUpdate(col, row, h, datos);
        if (val !== undefined) row[ci] = val;
      });
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { ok: true };
    }
  }
  return { ok: false };
}

// Elimina un registro por id.
function eliminarRecurso(tipo, id) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const conf  = _bdConf(tipo);
  const sheet = ss.getSheetByName(conf.hoja);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

// Configuración por tipo: qué hoja, cómo construir fila nueva y cómo actualizar.
function _bdConf(tipo) {
  const fechaHoy = () =>
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  if (tipo === "MATERIAL") {
    const calcCampos = d => ({
      codigo:              d.codigo            || "",
      categoria:           d.categoria         || "",
      nombre:              d.nombre            || "",
      unidad:              d.unidad            || "",
      precio_sin_iva:      parseFloat(d.precio_sin_iva) || 0,
      precio_con_iva:      Math.round((parseFloat(d.precio_sin_iva) || 0) * 1.19),
      precio_2026:         d.precio_2026       || "",
      proveedor:           d.proveedor         || "",
      fecha_actualizacion: fechaHoy(),
      partida:             d.partida           || "",
    });
    return {
      hoja: "Materiales",
      campoValor: (col, id, d) => col === "id" ? id : (calcCampos(d)[col] ?? ""),
      campoValorUpdate: (col, _row, _h, d) => {
        if (col === "id") return undefined;
        const m = calcCampos(d);
        return m[col] !== undefined ? m[col] : undefined;
      },
    };
  }

  if (tipo === "MANO_OBRA") {
    const calcDia = (sal, pct) =>
      Math.round((parseFloat(sal) || 0) * (1 + (parseFloat(pct) || 54) / 100) / 30);
    return {
      hoja: "ManoObra",
      campoValor: (col, id, d) => {
        if (col === "id")               return id;
        if (col === "descripcion")      return d.descripcion      || "";
        if (col === "salario_mensual")  return parseFloat(d.salario_mensual)  || 0;
        if (col === "prestaciones_pct") return parseFloat(d.prestaciones_pct) || 54;
        if (col === "costo_dia")        return calcDia(d.salario_mensual, d.prestaciones_pct);
        return "";
      },
      campoValorUpdate: (col, row, h, d) => {
        if (col === "id") return undefined;
        const sal = d.salario_mensual  !== undefined ? d.salario_mensual  : row[h.indexOf("salario_mensual")];
        const pct = d.prestaciones_pct !== undefined ? d.prestaciones_pct : row[h.indexOf("prestaciones_pct")];
        if (col === "descripcion")      return d.descripcion      !== undefined ? d.descripcion                   : undefined;
        if (col === "salario_mensual")  return d.salario_mensual  !== undefined ? parseFloat(d.salario_mensual)  : undefined;
        if (col === "prestaciones_pct") return d.prestaciones_pct !== undefined ? parseFloat(d.prestaciones_pct) : undefined;
        if (col === "costo_dia")        return calcDia(sal, pct);
        return undefined;
      },
    };
  }

  if (tipo === "EQUIPO") {
    return {
      hoja: "Equipos",
      campoValor: (col, id, d) => {
        if (col === "id")         return id;
        if (col === "nombre")     return d.nombre    || "";
        if (col === "tarifa_dia") return parseFloat(d.tarifa_dia) || 0;
        if (col === "partida")    return d.partida   || "";
        return "";
      },
      campoValorUpdate: (col, _row, _h, d) => {
        if (col === "id")         return undefined;
        if (col === "nombre"     && d.nombre     !== undefined) return d.nombre;
        if (col === "tarifa_dia" && d.tarifa_dia !== undefined) return parseFloat(d.tarifa_dia) || 0;
        if (col === "partida"    && d.partida    !== undefined) return d.partida;
        return undefined;
      },
    };
  }

  throw new Error("Tipo desconocido en _bdConf: " + tipo);
}
