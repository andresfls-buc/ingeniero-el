// ─── LISTA FIJA DE PARTIDAS ───────────────────────────────────────────────────

const PARTIDAS_FIJAS = [
  "Suministro Unidades Exteriores",
  "Suministro Unidades Interiores",
  "Tuberías de Refrigeración y Refrigerantes - Sistemas SRV",
  "Accesorios para Redes de Refrigeración - Sistemas SRV",
  "Aislamiento Térmico",
  "Redes de Distribución de Gas",
  "Mano de Obra e Instalación",
  "Transporte y Logística",
  "Otros",
];

// Mapeo: categoria de Materiales → partida
const CAT_A_PARTIDA = {
  "Accesorio CU":        "Accesorios para Redes de Refrigeración - Sistemas SRV",
  "Tub CU Flex ACR":     "Tuberías de Refrigeración y Refrigerantes - Sistemas SRV",
  "Tub CU Rig K":        "Tuberías de Refrigeración y Refrigerantes - Sistemas SRV",
  "Tub CU Rig L":        "Tuberías de Refrigeración y Refrigerantes - Sistemas SRV",
  "Tub Acero":           "Redes de Distribución de Gas",
  "Aislamiento":         "Aislamiento Térmico",
  "Tub Galvanizada Gas": "Redes de Distribución de Gas",
  "Tub PE-AL-PE Gas":    "Redes de Distribución de Gas",
};

function getPartidas() {
  return PARTIDAS_FIJAS;
}

function actualizarPartidaEquipo(equipoId, partida) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Equipos");
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];
  const idIdx  = h.indexOf("id");
  const parIdx = h.indexOf("partida");
  if (parIdx < 0) return { ok: false, error: "Columna 'partida' no existe en Equipos. Corre migrarPartidas() primero." };
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(equipoId)) {
      sheet.getRange(i + 1, parIdx + 1).setValue(partida);
      return { ok: true };
    }
  }
  return { ok: false, error: "Equipo no encontrado" };
}

// Devuelve la partida de un recurso según su tipo e ID en la BD.
function getPartidaRecurso(ss, tipo, recursoId) {
  if (tipo === "MANO_OBRA") return "Mano de Obra e Instalación";
  if (tipo === "OTRO")      return "Otros";
  if (!recursoId)           return "Otros";

  if (tipo === "MATERIAL") {
    const sheet = ss.getSheetByName("Materiales");
    if (!sheet) return "Otros";
    const data   = sheet.getDataRange().getValues();
    const h      = data[0];
    const idIdx  = h.indexOf("id");
    const catIdx = h.indexOf("categoria");
    const parIdx = h.indexOf("partida");
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(recursoId)) {
        if (parIdx >= 0 && data[i][parIdx]) return String(data[i][parIdx]);
        if (catIdx >= 0) return CAT_A_PARTIDA[String(data[i][catIdx]).trim()] || "Otros";
        return "Otros";
      }
    }
  }

  if (tipo === "EQUIPO") {
    const sheet = ss.getSheetByName("Equipos");
    if (!sheet) return "Suministro Unidades Exteriores";
    const data   = sheet.getDataRange().getValues();
    const h      = data[0];
    const idIdx  = h.indexOf("id");
    const parIdx = h.indexOf("partida");
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(recursoId)) {
        return (parIdx >= 0 && data[i][parIdx]) ? String(data[i][parIdx]) : "Suministro Unidades Exteriores";
      }
    }
  }

  return "Otros";
}

// ─── MIGRACIÓN ────────────────────────────────────────────────────────────────
// Agrega columna 'partida' a Materiales, Equipos y APU_Items en hojas existentes.

function migrarPartidas() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const msgs = [];

  // ── Materiales ────────────────────────────────────────────────────────────
  const matSheet = ss.getSheetByName("Materiales");
  if (matSheet) {
    const lastCol = matSheet.getLastColumn();
    const hRow = matSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    if (!hRow.includes("partida")) {
      const parCol = lastCol + 1;
      matSheet.getRange(1, parCol).setValue("partida");
      const data   = matSheet.getDataRange().getValues();
      const h      = data[0];
      const catIdx = h.indexOf("categoria");
      for (let i = 1; i < data.length; i++) {
        const cat = String(data[i][catIdx] || "").trim();
        matSheet.getRange(i + 1, parCol).setValue(CAT_A_PARTIDA[cat] || "Otros");
      }
      msgs.push("✅ Materiales: columna 'partida' agregada y mapeada desde categoría.");
    } else {
      msgs.push("— Materiales: 'partida' ya existía.");
    }
  }

  // ── Equipos ───────────────────────────────────────────────────────────────
  const eqSheet = ss.getSheetByName("Equipos");
  if (eqSheet) {
    const lastCol = eqSheet.getLastColumn();
    const hRow = eqSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    if (!hRow.includes("partida")) {
      eqSheet.getRange(1, lastCol + 1).setValue("partida");
      msgs.push("✅ Equipos: columna 'partida' agregada.\n   → Llena manualmente cada equipo con su partida (ej: 'Suministro Unidades Exteriores').");
    } else {
      msgs.push("— Equipos: 'partida' ya existía.");
    }
  }

  // ── APU_Items ─────────────────────────────────────────────────────────────
  const itemsSheet = ss.getSheetByName("APU_Items");
  if (itemsSheet && itemsSheet.getLastRow() > 0) {
    const lastCol = itemsSheet.getLastColumn();
    const hRow = itemsSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    if (!hRow.includes("partida")) {
      const parCol = lastCol + 1;
      itemsSheet.getRange(1, parCol).setValue("partida");
      if (itemsSheet.getLastRow() > 1) {
        const data    = itemsSheet.getDataRange().getValues();
        const h       = data[0];
        const tipoIdx = h.indexOf("tipo");
        const recIdx  = h.indexOf("recurso_id");
        for (let i = 1; i < data.length; i++) {
          const p = getPartidaRecurso(ss, data[i][tipoIdx], data[i][recIdx]);
          itemsSheet.getRange(i + 1, parCol).setValue(p);
        }
      }
      msgs.push("✅ APU_Items: columna 'partida' agregada y retroactivamente mapeada.");
    } else {
      msgs.push("— APU_Items: 'partida' ya existía.");
    }
  }

  SpreadsheetApp.getUi().alert(
    "Migración de Partidas completada\n\n" + msgs.join("\n")
  );
}
