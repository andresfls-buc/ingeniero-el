const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const code = fs.readFileSync(path.join(__dirname, '../src/apuHelpers.gs'), 'utf8');
const sandbox = {};
vm.runInNewContext(code + '\nthis.siguienteCodigoCopia = siguienteCodigoCopia;', sandbox);
const { siguienteCodigoCopia } = sandbox;

test('primera copia agrega "(copia)"', () => {
  assert.strictEqual(siguienteCodigoCopia('INS-01', ['INS-01']), 'INS-01 (copia)');
});

test('segunda copia agrega "(copia 2)"', () => {
  assert.strictEqual(
    siguienteCodigoCopia('INS-01', ['INS-01', 'INS-01 (copia)']),
    'INS-01 (copia 2)'
  );
});

test('tercera copia agrega "(copia 3)"', () => {
  assert.strictEqual(
    siguienteCodigoCopia('INS-01', ['INS-01', 'INS-01 (copia)', 'INS-01 (copia 2)']),
    'INS-01 (copia 3)'
  );
});

test('duplicar una copia parte del código base, no acumula "(copia) (copia)"', () => {
  assert.strictEqual(
    siguienteCodigoCopia('INS-01 (copia)', ['INS-01', 'INS-01 (copia)']),
    'INS-01 (copia 2)'
  );
});

test('código vacío produce solo "(copia)"', () => {
  assert.strictEqual(siguienteCodigoCopia('', []), '(copia)');
});
