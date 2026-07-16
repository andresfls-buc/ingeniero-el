// ─── Helpers puros de APU (sin SpreadsheetApp a nivel raíz — testeables con node) ───

/**
 * Devuelve el siguiente codigo_item de copia libre.
 * Quita cualquier sufijo "(copia)" o "(copia N)" del origen antes de numerar,
 * así "duplicar una copia" no acumula "(copia) (copia)".
 * El marcador vive SOLO en codigo_item; nunca llega al documento del cliente.
 */
function siguienteCodigoCopia(codigoOrigen, codigosExistentes) {
  var reSuf = /\s*\(copia(?:\s+\d+)?\)\s*$/i;
  var base  = String(codigoOrigen || "").replace(reSuf, "").trim();

  var existentes = {};
  (codigosExistentes || []).forEach(function (c) {
    existentes[String(c == null ? "" : c).trim()] = true;
  });

  function armar(n) {
    var suf = n === 1 ? "(copia)" : "(copia " + n + ")";
    return base ? base + " " + suf : suf;
  }

  var n = 1;
  while (existentes[armar(n)]) n++;
  return armar(n);
}
