# Duplicar APU — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Duplicar" button that copies an existing APU (header + all line items) into a new independent APU, marked only in the internal `codigo_item` so the marker never reaches a client document.

**Architecture:** One pure helper computes the client-safe copy code (`base (copia)`, `(copia 2)`, …) and is unit-tested with node. A backend `duplicarAPU(apuId)` reuses the already lock-protected `crearAPU` + `agregarItemAPU` to build the copy (no new id-generation code). The frontend adds a button in the APU list that calls it and opens the copy in the editor.

**Tech Stack:** Google Apps Script (`.gs`), Google Sheets, vanilla HTML/JS sidebar (`index.html`), node `--test` for pure helpers.

## Global Constraints

- Precios/valores en pesos colombianos (COP), sin decimales.
- El marcador de copia vive **solo en `codigo_item`**. `descripcion`, `actividad`, `unidad`, `cliente`, `direccion` de la copia se guardan **limpios**. El cliente ve `apu.actividad` en la cotización (cotizacion.gs:814) — nunca "(copia)".
- Toda escritura de fila que genera `id` debe pasar por `withLock` (patrón existente). Reusar `crearAPU` y `agregarItemAPU` cumple esto — NO escribir un copiador de filas nuevo.
- El usuario hace `clasp push` él mismo; nunca ejecutarlo como herramienta.
- Tests puros: el `.gs` cargado por node NO debe tocar `SpreadsheetApp` a nivel raíz (patrón de `cotizacionHelpers.gs`).

---

### Task 1: Helper puro `siguienteCodigoCopia`

**Files:**
- Create: `src/apuHelpers.gs`
- Test: `tests/apuHelpers.test.js`

**Interfaces:**
- Consumes: nada.
- Produces: `siguienteCodigoCopia(codigoOrigen, codigosExistentes) -> string` — dado el código del APU origen y un arreglo de todos los `codigo_item` existentes, devuelve el siguiente código de copia libre. Quita cualquier sufijo `(copia)` / `(copia N)` del origen antes de numerar.

- [ ] **Step 1: Write the failing test**

Create `tests/apuHelpers.test.js`:

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test`
Expected: FAIL — no puede leer `src/apuHelpers.gs` (no existe) o `siguienteCodigoCopia is not a function`.

- [ ] **Step 3: Write minimal implementation**

Create `src/apuHelpers.gs`:

```javascript
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test`
Expected: PASS — los 5 casos de `apuHelpers.test.js` pasan, y los tests existentes siguen pasando.

- [ ] **Step 5: Commit**

```bash
git add src/apuHelpers.gs tests/apuHelpers.test.js
git commit -m "feat: helper puro siguienteCodigoCopia para duplicar APU"
```

---

### Task 2: Backend `duplicarAPU(apuId)`

**Files:**
- Modify: `src/apu.gs` (agregar función nueva al final, junto a las demás funciones de APU)

**Interfaces:**
- Consumes:
  - `siguienteCodigoCopia(codigoOrigen, codigosExistentes)` (Task 1).
  - `getAPUCompleto(apuId)` → devuelve `{ id, codigo_item, unidad, cliente, direccion, actividad, administracion_pct, imprevistos_pct, utilidad_pct, iva_pct, desperdicio_pct, herramienta_menor_pct, equipos:[], materiales:[], mano_obra:[], otros:[] }`; cada ítem tiene `{ tipo, recurso_id, descripcion_manual, cantidad, rendimiento, precio_unitario }` (apu.gs:710).
  - `crearAPU(datos)` → `newId` (lock-protegido, subtotales en 0) (apu.gs:620).
  - `actualizarCabezaAPU(apuId, datos)` (apu.gs:664).
  - `agregarItemAPU(apuId, item)` (lock-protegido, recalcula valor_parcial + subtotales) (apu.gs:766).
- Produces: `duplicarAPU(apuId) -> { id: number }` — llamable desde el frontend vía `google.script.run`.

- [ ] **Step 1: Escribir la implementación**

Agregar al final de `src/apu.gs`:

```javascript
// ─── DUPLICAR APU ─────────────────────────────────────────────────────────────
// Crea un APU independiente copiando cabeza + ítems del origen.
// El marcador "(copia)" vive SOLO en codigo_item (interno). La descripción se guarda
// LIMPIA para que el fallback de la cotización (actividad || descripcion || codigo_item)
// nunca muestre el marcador al cliente.
function duplicarAPU(apuId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = getAPUCompleto(apuId);
  if (!src) throw new Error("APU no encontrado: " + apuId);

  // Código de la copia: leer todos los codigo_item existentes y pedir el siguiente libre.
  const apuSheet = ss.getSheetByName("APU");
  const apuData  = apuSheet.getDataRange().getValues();
  const colCod   = apuData[0].indexOf("codigo_item");
  const codigos  = apuData.slice(1).map(function (r) { return r[colCod]; });
  const nuevoCodigo = siguienteCodigoCopia(src.codigo_item, codigos);

  // Descripción limpia (código base, sin "(copia)").
  const descLimpia = String(src.codigo_item || "")
    .replace(/\s*\(copia(?:\s+\d+)?\)\s*$/i, "").trim();

  // 1) Cabeza (crearAPU arranca subtotales en 0).
  const newId = crearAPU({
    codigo_item: nuevoCodigo,
    descripcion: descLimpia,
    unidad:      src.unidad,
    cliente:     src.cliente,
    direccion:   src.direccion,
    actividad:   src.actividad,
  });

  // 2) AIU / desperdicio / HM.
  actualizarCabezaAPU(newId, {
    administracion_pct:    src.administracion_pct,
    imprevistos_pct:       src.imprevistos_pct,
    utilidad_pct:          src.utilidad_pct,
    iva_pct:               src.iva_pct,
    desperdicio_pct:       src.desperdicio_pct,
    herramienta_menor_pct: src.herramienta_menor_pct,
  });

  // 3) Copiar ítems. El último agregarItemAPU recalcula subtotales + costo_neto + valor_total.
  const items = []
    .concat(src.equipos    || [])
    .concat(src.materiales || [])
    .concat(src.mano_obra  || [])
    .concat(src.otros      || []);

  items.forEach(function (it) {
    agregarItemAPU(newId, {
      tipo:               it.tipo,
      recurso_id:         it.recurso_id,
      descripcion_manual: it.descripcion_manual,
      cantidad:           it.cantidad,
      rendimiento:        it.rendimiento,
      precio_unitario:    it.precio_unitario,
    });
  });

  return { id: newId };
}
```

- [ ] **Step 2: Verificar sintaxis**

Run: `node --check src/apu.gs`
Expected: sin salida (exit 0). Si falla, corregir el error de sintaxis antes de seguir.

- [ ] **Step 3: Commit**

```bash
git add src/apu.gs
git commit -m "feat: duplicarAPU backend (reusa crearAPU + agregarItemAPU, lock-safe)"
```

---

### Task 3: Botón "Duplicar" en la lista de APUs (frontend)

**Files:**
- Modify: `src/index.html` — `renderLista` (~1378-1401) y agregar `duplicarAPUenLista` cerca de `abrirAPU` (~1415).

**Interfaces:**
- Consumes: `duplicarAPU(apuId)` del servidor (Task 2) vía `google.script.run`; `abrirAPU(id)`, `setEstado(msg)` existentes.
- Produces: interacción de usuario; sin API nueva.

- [ ] **Step 1: Agregar el botón en la fila de la lista**

En `src/index.html`, dentro de `renderLista`, reemplazar la celda de acciones:

```html
        <td style="display:flex;gap:6px">
          <button class="btn-open" onclick="abrirAPU(${a.id})">Abrir →</button>
          <button class="btn-del" title="Eliminar APU" onclick="confirmarEliminarAPU(${a.id}, '${String(a.cliente || a.descripcion || '').replace(/'/g,"\\'")}')">✕</button>
        </td>
```

por (agrega el botón "⧉ Duplicar" entre Abrir y ✕):

```html
        <td style="display:flex;gap:6px">
          <button class="btn-open" onclick="abrirAPU(${a.id})">Abrir →</button>
          <button class="btn-dup" title="Duplicar APU" onclick="duplicarAPUenLista(this, ${a.id})">⧉ Duplicar</button>
          <button class="btn-del" title="Eliminar APU" onclick="confirmarEliminarAPU(${a.id}, '${String(a.cliente || a.descripcion || '').replace(/'/g,"\\'")}')">✕</button>
        </td>
```

- [ ] **Step 2: Agregar la función `duplicarAPUenLista`**

En `src/index.html`, justo después de la función `abrirAPU` (~línea 1429), agregar:

```javascript
  function duplicarAPUenLista(btn, id) {
    // Feedback visible inmediato en el propio botón.
    const textoOrig = btn.textContent;
    btn.disabled = true;
    btn.textContent = "⏳ Duplicando...";
    google.script.run
      .withSuccessHandler(function (res) {
        if (!res || !res.id) {
          btn.disabled = false;
          btn.textContent = textoOrig;
          alert("No se pudo duplicar el APU.");
          return;
        }
        // Abre la copia en el editor y confirma.
        abrirAPU(res.id);
        setEstado("APU duplicado ✓ — edita la copia");
      })
      .withFailureHandler(function (err) {
        btn.disabled = false;
        btn.textContent = textoOrig;
        alert("Error al duplicar: " + err.message);
      })
      .duplicarAPU(id);
  }
```

- [ ] **Step 3: (Opcional) estilo del botón**

Si existe una hoja de estilos con `.btn-open` / `.btn-del`, agregar una regla `.btn-dup` con el mismo tamaño/tipografía y un color neutro (p. ej. reutilizar el estilo de `.btn-open`). Si no hay estilos dedicados (los botones heredan estilos base), omitir este paso — el botón funciona igual.

- [ ] **Step 4: Verificación manual (la hace el usuario)**

El usuario hace `clasp push` y en el sidebar:
1. En la lista de APUs, clic en "⧉ Duplicar" de un APU con varios ítems → aparece la copia **abierta en el editor**, con estado "APU duplicado ✓", con los mismos ítems, subtotales y % de AIU que el origen.
2. Volver a la lista: la copia aparece con código `CÓDIGO (copia)`. Duplicar el mismo origen otra vez → `CÓDIGO (copia 2)`. Duplicar la copia → `CÓDIGO (copia 2)` o siguiente libre (parte del código base).
3. Agregar la copia a una **cotización cliente** → el documento del cliente muestra la **actividad** y su valor, **sin** "(copia)".

- [ ] **Step 5: Commit**

```bash
git add src/index.html
git commit -m "feat: botón Duplicar en la lista de APUs (abre la copia para editar)"
```

---

## Self-Review

**Spec coverage:**
- Marcador solo en `codigo_item`, descripción limpia → Task 2 (crearAPU con `descripcion: descLimpia`). ✓
- No leak al cliente (usa actividad) → verificado en Task 3 Step 4.3. ✓
- Numeración `(copia)`, `(copia 2)`, … + duplicar una copia sin acumular → Task 1 (5 casos). ✓
- Copia independiente, seleccionable en cotizaciones → Task 2 crea APU real; verificado en Task 3 Step 4.3. ✓
- Reusar `crearAPU` + `agregarItemAPU` (lock-safe) → Task 2. ✓
- Botón visible + abre la copia + confirmación → Task 3. ✓
- Sin columnas nuevas / sin migración → ningún task toca setup.gs. ✓

**Placeholder scan:** Sin TBD/TODO. Task 3 Step 3 es explícitamente opcional y describe la condición para omitirlo. ✓

**Type consistency:** `siguienteCodigoCopia(codigoOrigen, codigosExistentes)` usado con la misma firma en Task 1 y Task 2. `duplicarAPU(apuId) -> {id}` producido en Task 2 y consumido en Task 3. Campos de ítem (`tipo, recurso_id, descripcion_manual, cantidad, rendimiento, precio_unitario`) coinciden entre `getAPUCompleto` y `agregarItemAPU`. ✓
