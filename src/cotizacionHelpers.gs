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
