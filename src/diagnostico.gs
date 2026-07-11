// ─── DIAGNÓSTICO TEMPORAL (SOLO LECTURA) ─────────────────────────────────────
// Revisa las cotizaciones ya guardadas y reporta señales del bug de IDs
// duplicados (ítems que se pisan / desaparecen al descargar). NO modifica nada.
// Borrar este archivo cuando ya no se necesite.

function diagnosticarCotizaciones() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = ss.getSheetByName("Cotizacion_Items");
  const cotSheet   = ss.getSheetByName("Cotizaciones");
  if (!itemsSheet || !cotSheet) return "No se encontraron las hojas de cotización.";

  const idata = itemsSheet.getDataRange().getValues();
  const cdata = cotSheet.getDataRange().getValues();
  if (idata.length < 2) return "No hay ítems de cotización guardados.";

  const ih     = idata[0];
  const idIdx  = 0;
  const cIdx   = ih.indexOf("cotizacion_id");
  const numIdx = ih.indexOf("item_num");
  const desIdx = ih.indexOf("descripcion");

  // Nombre legible de cada cotización
  const ch = cdata[0];
  const nOfIdx = ch.indexOf("numero_oferta");
  const cliIdx = ch.indexOf("cliente");
  const nombreCot = {};
  for (let i = 1; i < cdata.length; i++) {
    nombreCot[String(cdata[i][0])] =
      (cdata[i][nOfIdx] || ("cot " + cdata[i][0])) +
      (cdata[i][cliIdx] ? " — " + cdata[i][cliIdx] : "");
  }

  // Agrupar ítems por cotización
  const porCot = {};
  for (let i = 1; i < idata.length; i++) {
    const cot = String(idata[i][cIdx]);
    if (!porCot[cot]) porCot[cot] = [];
    porCot[cot].push({
      id:       String(idata[i][idIdx]),
      item_num: String(idata[i][numIdx] == null ? "" : idata[i][numIdx]).trim(),
      desc:     String(idata[i][desIdx] || "").slice(0, 30),
    });
  }

  const repetidos = arr => {
    const vistos = {}, dup = {};
    arr.forEach(v => { if (vistos[v]) dup[v] = true; else vistos[v] = true; });
    return Object.keys(dup);
  };

  let dañadas = 0;
  const lineas = [];
  Object.keys(porCot).forEach(cot => {
    const items   = porCot[cot];
    const dupIds  = repetidos(items.map(x => x.id));
    const dupNums = repetidos(items.map(x => x.item_num).filter(x => x !== ""));
    const problema = dupIds.length > 0;   // IDs repetidos = daño real
    if (problema) dañadas++;

    lineas.push(
      (problema ? "⚠️ DAÑADA  " : "✅ OK      ") +
      (nombreCot[cot] || ("cot " + cot)) +
      "\n     ítems: " + items.length +
      "  |  IDs repetidos: " + (dupIds.length ? dupIds.join(", ") : "no") +
      "  |  números repetidos: " + (dupNums.length ? dupNums.join(", ") : "no")
    );
  });

  const encabezado =
    "DIAGNÓSTICO DE COTIZACIONES\n" +
    "Cotizaciones revisadas: " + Object.keys(porCot).length +
    "  |  con daño (IDs repetidos): " + dañadas + "\n" +
    "──────────────────────────────────────────\n";

  const reporte = encabezado + lineas.sort().join("\n\n");
  Logger.log(reporte);
  return reporte;
}
