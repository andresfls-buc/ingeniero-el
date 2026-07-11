// ─── HELPERS PUROS DE COTIZACIÓN (testeables con Node) ────────────────────────
// Sin dependencias de SpreadsheetApp. Usados por el renderer del cliente.

// Compara dos números de ítem jerárquicos ("4.1", "10.2") por segmentos numéricos.
// "4.2" < "10.1" (4 < 10). Los vacíos/null van al final.
function compararItemNum(a, b) {
  var sa = String(a == null ? "" : a).trim();
  var sb = String(b == null ? "" : b).trim();
  if (!sa && !sb) return 0;
  if (!sa) return 1;
  if (!sb) return -1;
  var pa = sa.split(".");
  var pb = sb.split(".");
  var n = Math.max(pa.length, pb.length);
  for (var i = 0; i < n; i++) {
    var na = parseFloat(pa[i]); var nb = parseFloat(pb[i]);
    var va = isNaN(na) ? 0 : na; var vb = isNaN(nb) ? 0 : nb;
    if (va !== vb) return va - vb;
  }
  return 0;
}

// Fórmula ÚNICA de totales de una cotización. La usan el servidor
// (recalcularCotizacion), el documento del cliente (llenarHojaCotizacionCliente)
// y el frontend (recalcularTotalCot) para que lista, web y PDF muestren SIEMPRE
// el mismo número.
//
// Regla (igual a la del documento que recibe el cliente): se redondea CADA
// componente del AIU antes de sumarlos, luego el IVA se aplica sobre neto+AIU.
// pcts: { administracion_pct, imprevistos_pct, utilidad_pct, iva_pct }
function calcularTotalesCotizacion(neto, pcts) {
  pcts = pcts || {};
  var n       = parseFloat(neto) || 0;
  var admPct  = parseFloat(pcts.administracion_pct) || 0;
  var impPct  = parseFloat(pcts.imprevistos_pct)    || 0;
  var utilPct = parseFloat(pcts.utilidad_pct)       || 0;
  var ivaPct  = parseFloat(pcts.iva_pct)            || 0;

  var admVal  = Math.round(n * admPct  / 100);
  var impVal  = Math.round(n * impPct  / 100);
  var utilVal = Math.round(n * utilPct / 100);
  var sinIVA  = n + admVal + impVal + utilVal;
  var ivaVal  = Math.round(sinIVA * ivaPct / 100);
  var total   = sinIVA + ivaVal;

  return {
    neto:    Math.round(n),
    admVal:  admVal,
    impVal:  impVal,
    utilVal: utilVal,
    sinIVA:  sinIVA,
    ivaVal:  ivaVal,
    total:   total,
  };
}
