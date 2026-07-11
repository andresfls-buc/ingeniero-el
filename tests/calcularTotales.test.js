const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const code = fs.readFileSync(path.join(__dirname, '../src/cotizacionHelpers.gs'), 'utf8');
const sandbox = {};
vm.runInNewContext(code + '\nthis.calcularTotalesCotizacion = calcularTotalesCotizacion;', sandbox);
const { calcularTotalesCotizacion } = sandbox;

test('sin AIU ni IVA: total == neto', () => {
  const r = calcularTotalesCotizacion(1000000, {});
  assert.strictEqual(r.neto, 1000000);
  assert.strictEqual(r.total, 1000000);
});

test('redondea cada componente AIU antes de sumar (igual que el PDF del cliente)', () => {
  // neto que produce decimales en los componentes
  const neto = 1234567;
  const r = calcularTotalesCotizacion(neto, {
    administracion_pct: 5, imprevistos_pct: 3, utilidad_pct: 7, iva_pct: 19,
  });
  const admVal  = Math.round(neto * 5 / 100);
  const impVal  = Math.round(neto * 3 / 100);
  const utilVal = Math.round(neto * 7 / 100);
  const sinIVA  = neto + admVal + impVal + utilVal;
  const ivaVal  = Math.round(sinIVA * 19 / 100);
  assert.strictEqual(r.admVal, admVal);
  assert.strictEqual(r.impVal, impVal);
  assert.strictEqual(r.utilVal, utilVal);
  assert.strictEqual(r.sinIVA, sinIVA);
  assert.strictEqual(r.ivaVal, ivaVal);
  assert.strictEqual(r.total, sinIVA + ivaVal);
});

test('IVA se aplica sobre neto+AIU, no solo sobre neto', () => {
  const r = calcularTotalesCotizacion(1000000, { utilidad_pct: 10, iva_pct: 19 });
  // sinIVA = 1.100.000 ; iva = 209.000 ; total = 1.309.000
  assert.strictEqual(r.sinIVA, 1100000);
  assert.strictEqual(r.ivaVal, 209000);
  assert.strictEqual(r.total, 1309000);
});

test('acepta porcentajes como string o number y neto inválido → 0', () => {
  const r = calcularTotalesCotizacion('abc', { utilidad_pct: '10' });
  assert.strictEqual(r.neto, 0);
  assert.strictEqual(r.total, 0);
});

test('el total del servidor coincide con el del documento del cliente (mismo redondeo)', () => {
  // Regresión del bug de dinero: antes recalcularCotizacion multiplicaba todo y
  // redondeaba una sola vez, dando un valor distinto al del PDF.
  const neto = 987654;
  const pcts = { administracion_pct: 8, imprevistos_pct: 2, utilidad_pct: 5, iva_pct: 19 };
  const r = calcularTotalesCotizacion(neto, pcts);

  // Reproducción EXACTA del render del cliente (llenarHojaCotizacionCliente)
  const admVal  = Math.round(neto * 8 / 100);
  const impVal  = Math.round(neto * 2 / 100);
  const utilVal = Math.round(neto * 5 / 100);
  const sinIVA  = neto + admVal + impVal + utilVal;
  const ivaVal  = Math.round(sinIVA * 19 / 100);
  const totalCliente = sinIVA + ivaVal;

  assert.strictEqual(r.total, totalCliente);
});
