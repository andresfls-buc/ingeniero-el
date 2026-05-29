const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// Cargar el .gs en un sandbox (sus funciones no tocan SpreadsheetApp a nivel raíz)
const code = fs.readFileSync(path.join(__dirname, '../src/cotizacionHelpers.gs'), 'utf8');
const sandbox = {};
vm.runInNewContext(code + '\nthis.compararItemNum = compararItemNum;', sandbox);
const { compararItemNum } = sandbox;

test('ordena numéricamente, no alfabéticamente (4.2 antes de 10.1)', () => {
  assert.ok(compararItemNum('4.2', '10.1') < 0);
  assert.ok(compararItemNum('10.1', '4.2') > 0);
});

test('ordena una lista completa por segmentos numéricos', () => {
  const arr = ['10.2', '4.1', '4.2', '5.0', '1.1'];
  arr.sort(compararItemNum);
  assert.deepStrictEqual(arr, ['1.1', '4.1', '4.2', '5.0', '10.2']);
});

test('los ítems sin número van al final', () => {
  const arr = ['2.1', '', '1.1'];
  arr.sort(compararItemNum);
  assert.deepStrictEqual(arr, ['1.1', '2.1', '']);
});

test('iguales devuelven 0', () => {
  assert.strictEqual(compararItemNum('3.3', '3.3'), 0);
});

test('tolera null/undefined tratándolos como vacío', () => {
  assert.strictEqual(compararItemNum(null, undefined), 0);
  assert.ok(compararItemNum('1.1', null) < 0);
});
