# Cotización para el cliente (documento limpio) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generar, desde la misma cotización, un segundo documento limpio para el cliente final (sin desglose de equipos/materiales/MO), con número de ítem y descripción libres, líneas mixtas (APU o manuales) e ítems ordenados numéricamente, escalable a 300+ ítems.

**Architecture:** El documento interno de auditoría (`llenarHojaCotizacion`) queda intacto. Se agrega un renderer nuevo y autónomo (`llenarHojaCotizacionCliente`) que dibuja una tabla de 6 columnas en bloque (una sola `setValues` + formatos por `RangeList`) para no superar el límite de 6 min de Apps Script. Una lectura liviana (`getCotizacionCliente`) evita cargar `APU_Items`/BD. Los exports reciben un parámetro `tipo` ("cliente" | "interna").

**Tech Stack:** Google Apps Script (`.gs`), Google Sheets, HTML/JS sidebar (`index.html`), clasp para deploy. Pruebas unitarias de lógica pura con el runner integrado de Node (`node --test`) vía `vm`.

> **Nota de entorno:** El push a Apps Script lo hace el usuario manualmente. NUNCA ejecutes `clasp push` como herramienta; cuando un paso requiera código vivo en la hoja, pídele al usuario que corra `clasp push` y luego verifique.

> **Decisión de diseño documentada:** El renderer del cliente es autónomo (construye su propio encabezado/pie) en vez de compartir helpers con `llenarHojaCotizacion`. Es una desviación deliberada del spec ("extraer a helpers compartidos") para **no arriesgar una regresión visual en el documento interno de auditoría**, que debe quedar idéntico. Se acepta una duplicación acotada de encabezado/pie.

---

## File Structure

- **Create:** `src/cotizacionHelpers.gs` — helpers puros y testeables (orden de ítems). Sin dependencias de `SpreadsheetApp`.
- **Create:** `tests/cotizacionHelpers.test.js` — pruebas Node de la lógica pura.
- **Create:** `package.json` script `test` apuntando a `node --test`.
- **Modify:** `src/cotizacion.gs` — `getCotizacionCliente` (nueva), `llenarHojaCotizacionCliente` (nueva), `agregarLineaManual` (nueva), `agregarVariasLineas` (nueva), `actualizarItemCotizacion` (nueva, generaliza a `actualizarCantidadItem`), y parámetro `tipo` en exports.
- **Modify:** `src/index.html` — campos editables `item_num`/`descripcion`, modal "línea manual", dos botones de export.

---

## Task 1: Helper de orden de ítems (lógica pura, TDD)

**Files:**
- Create: `src/cotizacionHelpers.gs`
- Create: `tests/cotizacionHelpers.test.js`
- Modify: `package.json`

- [ ] **Step 1: Configurar el runner de pruebas en package.json**

Reemplazar la línea del script `test` en `package.json`:

```json
    "test": "node --test"
```

- [ ] **Step 2: Escribir el test que falla**

Crear `tests/cotizacionHelpers.test.js` con contenido completo:

```js
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
```

- [ ] **Step 3: Correr el test para verificar que falla**

Run: `npm test`
Expected: FALLA — `ENOENT` (no existe `src/cotizacionHelpers.gs`) o `compararItemNum is not a function`.

- [ ] **Step 4: Implementar el helper**

Crear `src/cotizacionHelpers.gs`:

```js
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
```

- [ ] **Step 5: Correr el test para verificar que pasa**

Run: `npm test`
Expected: PASA — 5 tests ok.

- [ ] **Step 6: Commit**

```bash
git add src/cotizacionHelpers.gs tests/cotizacionHelpers.test.js package.json
git commit -m "feat: helper compararItemNum para ordenar ítems de cotización"
```

---

## Task 2: Lectura liviana `getCotizacionCliente`

**Files:**
- Modify: `src/cotizacion.gs`

Lee solo `Cotizaciones` (1 fila) + `Cotizacion_Items` (filtrado). NO toca `APU_Items` ni la BD. Devuelve los ítems ya ordenados por `item_num`.

- [ ] **Step 1: Implementar la función**

Agregar al final de `src/cotizacion.gs`:

```js
// ─── GET COTIZACIÓN CLIENTE (liviana, escala a 300+ ítems) ───────────────────
// A diferencia de getCotizacionCompleta, NO carga APU_Items ni la BD de materiales.
// El documento del cliente no muestra desglose, así que no se necesitan sub-ítems.
function getCotizacionCliente(cotId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cotSheet = ss.getSheetByName("Cotizaciones");
  const cotData  = cotSheet.getDataRange().getValues();
  const cotH     = cotData[0];
  const cotRow   = cotData.slice(1).find(r => r[0] == cotId);
  if (!cotRow) return null;

  const cot = {};
  cotH.forEach((h, i) => cot[h] = cotRow[i]);

  // Normalizar % guardados como decimal (0.5 → 50)
  ["administracion_pct", "imprevistos_pct", "utilidad_pct", "iva_pct"].forEach(k => {
    const v = parseFloat(cot[k]) || 0;
    cot[k] = (v > 0 && v < 1) ? Math.round(v * 100) : v;
  });

  const itemsSheet = ss.getSheetByName("Cotizacion_Items");
  const itemsData  = itemsSheet.getDataRange().getValues();
  if (itemsData.length < 2) { cot.items = []; return cot; }

  const ih = itemsData[0];
  cot.items = itemsData.slice(1)
    .filter(r => r[ih.indexOf("cotizacion_id")] == cotId)
    .map(r => { const o = {}; ih.forEach((h, i) => o[h] = r[i]); return o; });

  // Ordenar por número de ítem (orden numérico jerárquico)
  cot.items.sort((a, b) => compararItemNum(a.item_num, b.item_num));

  return cot;
}
```

- [ ] **Step 2: Verificación manual (requiere push del usuario)**

Pídele al usuario: "Corre `clasp push` y luego, en el editor de Apps Script, ejecuta esta función de prueba y mira el log."

Pegar temporalmente en el editor y ejecutar:

```js
function _probarGetCotizacionCliente() {
  const cots = listarCotizaciones();
  if (!cots.length) { Logger.log("No hay cotizaciones para probar"); return; }
  const cot = getCotizacionCliente(cots[0].id);
  Logger.log("items: " + (cot.items || []).length);
  Logger.log("primer item_num: " + (cot.items[0] ? cot.items[0].item_num : "—"));
  Logger.log("tiene equipos?: " + (cot.items[0] && cot.items[0].equipos !== undefined));
}
```

Expected: loguea la cantidad de ítems; `tiene equipos?: false` (confirma que no carga desglose).

- [ ] **Step 3: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: getCotizacionCliente (lectura liviana, ítems ordenados)"
```

---

## Task 3: Editar `item_num` y `descripcion`; generalizar actualización de ítem

**Files:**
- Modify: `src/cotizacion.gs`

Hoy solo existe `actualizarCantidadItem`. Se agrega `actualizarItemCotizacion(itemId, cambios)` que permite editar también `item_num`, `descripcion`, `unidad` y `precio_apu` (estos dos últimos para líneas manuales), recalculando `valor_total` y los totales de la cotización.

- [ ] **Step 1: Implementar la función**

Agregar al final de `src/cotizacion.gs`:

```js
// ─── ACTUALIZAR ÍTEM DE COTIZACIÓN (campos múltiples) ────────────────────────
// Permite editar item_num, descripcion, unidad, cantidad y precio_apu.
// Recalcula valor_total (= cantidad * precio_apu) y los totales de la cotización.
function actualizarItemCotizacion(itemId, cambios) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const row   = [...data[i]];
      const cotId = row[h.indexOf("cotizacion_id")];

      ["item_num", "descripcion", "unidad"].forEach(campo => {
        if (cambios[campo] !== undefined) {
          const c = h.indexOf(campo);
          if (c >= 0) row[c] = cambios[campo];
        }
      });

      const cant   = parseFloat(cambios.cantidad   ?? row[h.indexOf("cantidad")])   || 0;
      const precio = parseFloat(cambios.precio_apu ?? row[h.indexOf("precio_apu")]) || 0;
      row[h.indexOf("cantidad")]    = cant;
      row[h.indexOf("precio_apu")]  = precio;
      row[h.indexOf("valor_total")] = cant * precio;

      sheet.getRange(i + 1, 1, 1, h.length).setValues([row]);
      recalcularCotizacion(ss, cotId);
      return { ok: true, valor_total: cant * precio };
    }
  }
  return { ok: false };
}
```

- [ ] **Step 2: Verificación manual (requiere push del usuario)**

Pídele al usuario que corra `clasp push`. Luego, en el editor:

```js
function _probarActualizarItem() {
  const cots = listarCotizaciones();
  const cot  = getCotizacionCliente(cots[0].id);
  const it   = cot.items[0];
  Logger.log("antes: num=" + it.item_num + " desc=" + it.descripcion);
  actualizarItemCotizacion(it.id, { item_num: "9.9", descripcion: "PRUEBA EDIT" });
  const cot2 = getCotizacionCliente(cots[0].id);
  const it2  = cot2.items.find(x => x.id == it.id);
  Logger.log("después: num=" + it2.item_num + " desc=" + it2.descripcion);
}
```

Expected: el log muestra `después: num=9.9 desc=PRUEBA EDIT`.

- [ ] **Step 3: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: actualizarItemCotizacion (editar item_num, descripcion, unidad, precio)"
```

---

## Task 4: Líneas manuales — individual y por lote

**Files:**
- Modify: `src/cotizacion.gs`

`agregarLineaManual` inserta una fila con `apu_id` vacío. `agregarVariasLineas` inserta N filas en una sola escritura (escalabilidad para licitaciones grandes), recalculando totales una sola vez.

- [ ] **Step 1: Implementar las dos funciones**

Agregar al final de `src/cotizacion.gs`:

```js
// ─── AGREGAR LÍNEA MANUAL (sin APU) ──────────────────────────────────────────
function agregarLineaManual(cotId, datos) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName("Cotizacion_Items");
  const data   = sheet.getDataRange().getValues();
  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId  = lastId + 1;

  const cant   = parseFloat(datos.cantidad)   || 0;
  const precio = parseFloat(datos.precio_apu) || 0;

  // Orden de columnas: id, cotizacion_id, apu_id, item_num, descripcion, unidad, cantidad, precio_apu, valor_total
  sheet.appendRow([
    newId, cotId, "",
    datos.item_num    || "",
    datos.descripcion || "",
    datos.unidad      || "",
    cant, precio, cant * precio
  ]);

  recalcularCotizacion(ss, cotId);
  return { id: newId, valor_total: cant * precio };
}

// ─── AGREGAR VARIAS LÍNEAS DE UN GOLPE (escala a 300+ ítems) ─────────────────
// lineas: [{ item_num, descripcion, unidad, cantidad, precio_apu }, ...]
// Una sola escritura a la hoja + un solo recálculo.
function agregarVariasLineas(cotId, lineas) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  let nextId  = (data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0) + 1;

  const filas = (lineas || []).map(l => {
    const cant   = parseFloat(l.cantidad)   || 0;
    const precio = parseFloat(l.precio_apu) || 0;
    return [
      nextId++, cotId, "",
      l.item_num    || "",
      l.descripcion || "",
      l.unidad      || "",
      cant, precio, cant * precio
    ];
  });
  if (!filas.length) return { ok: true, count: 0 };

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, filas.length, filas[0].length).setValues(filas);
  recalcularCotizacion(ss, cotId);
  return { ok: true, count: filas.length };
}
```

- [ ] **Step 2: Verificación manual (requiere push del usuario)**

Pídele al usuario que corra `clasp push`. Luego, en el editor:

```js
function _probarLineasManuales() {
  const cots = listarCotizaciones();
  const cotId = cots[0].id;
  agregarLineaManual(cotId, { item_num: "1.1", descripcion: "Localización y replanteo", unidad: "GL", cantidad: 1, precio_apu: 250000 });
  agregarVariasLineas(cotId, [
    { item_num: "1.2", descripcion: "Campamento", unidad: "GL", cantidad: 1, precio_apu: 500000 },
    { item_num: "2.1", descripcion: "Tubería 1\"", unidad: "ml", cantidad: 30, precio_apu: 18000 }
  ]);
  const cot = getCotizacionCliente(cotId);
  cot.items.forEach(i => Logger.log(i.item_num + " | " + i.descripcion + " | " + i.valor_total));
}
```

Expected: el log lista las 3 (o más) líneas; las nuevas tienen su `valor_total` correcto (250000, 500000, 540000) y salen ordenadas por número.

- [ ] **Step 3: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: agregarLineaManual y agregarVariasLineas (líneas sin APU, escritura por lote)"
```

---

## Task 5: Renderer del documento del cliente (batch)

**Files:**
- Modify: `src/cotizacion.gs`

Renderer autónomo. Construye la tabla de ítems en una matriz y la escribe con un solo `setValues`; los formatos se aplican por rango (no celda por celda). Encabezado/pie propios para no tocar `llenarHojaCotizacion`.

- [ ] **Step 1: Implementar la función**

Agregar al final de `src/cotizacion.gs`:

```js
// ─── RENDER DOCUMENTO DEL CLIENTE (tabla limpia, sin desglose) ───────────────
// Diseñado para 300+ ítems: la tabla de ítems se escribe con UN solo setValues
// y los formatos se aplican por rangos completos, no celda por celda.
function llenarHojaCotizacionCliente(sheet, cot) {
  const items   = (cot.items || []).slice().sort((a, b) => compararItemNum(a.item_num, b.item_num));
  const valorNeto = parseFloat(cot.valor_neto) || 0;
  const admPct  = parseFloat(cot.administracion_pct) || 0;
  const impPct  = parseFloat(cot.imprevistos_pct)    || 0;
  const utilPct = parseFloat(cot.utilidad_pct)       || 0;
  const ivaPct  = parseFloat(cot.iva_pct)            || 0;

  const cfg       = getConfig();
  const empresa   = (cfg["empresa"]          || "").trim();
  const remitente = (cfg["nombre_remitente"] || "").trim();
  const logoId    = (cfg["logo_id"]          || "").trim();
  const firmaId   = (cfg["firma_id"]         || "").trim();

  const formaPago     = String(cot.forma_pago     || "").trim();
  const plazoEntrega  = String(cot.plazo_entrega  || "").trim();
  const validezOferta = String(cot.validez_oferta || "").trim();
  const noIncluye     = String(cot.no_incluye     || "").trim();

  const MONEY = '"$"#,##0';
  const NEGRO = "#a8a8a8", GRIS_OSC = "#c8c8c8", BORDE = "#dddddd";
  const NC = 6;

  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 55);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 130);

  let r = 1;

  // Encabezado: logo + empresa
  sheet.setRowHeight(r, 55);
  sheet.getRange(r, 1, 1, 2).merge().setValue("").setBackground("#ffffff");
  sheet.getRange(r, 3, 1, 4).merge()
    .setValue(empresa || "").setFontSize(16).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#ffffff");
  if (logoId) {
    try { sheet.insertImage(DriveApp.getFileById(logoId).getBlob(), 1, r).setWidth(160).setHeight(48); } catch(e) {}
  }
  r++;

  // Cliente + dirección
  sheet.setRowHeight(r, 22);
  sheet.getRange(r, 1, 1, 3).merge().setValue(cot.cliente ? "CLIENTE:    " + cot.cliente : "")
    .setFontSize(9).setFontColor("#222222").setFontWeight("bold")
    .setVerticalAlignment("middle").setWrap(true).setBackground("#f5f5f5");
  sheet.getRange(r, 4, 1, 3).merge().setValue(cot.direccion ? "DIRECCIÓN:  " + cot.direccion : "")
    .setFontSize(9).setFontColor("#333333").setVerticalAlignment("middle").setWrap(true).setBackground("#f5f5f5");
  r++;

  // N° oferta + fecha
  sheet.setRowHeight(r, 14);
  const infoOferta = [
    cot.numero_oferta ? "N° Oferta: " + cot.numero_oferta : null,
    cot.fecha         ? "Fecha: "     + cot.fecha         : null,
  ].filter(Boolean).join("     ");
  sheet.getRange(r, 1, 1, NC).merge().setValue(infoOferta)
    .setFontSize(9).setFontColor("#555555").setHorizontalAlignment("right").setVerticalAlignment("middle");
  r++;

  // Encabezado de tabla
  sheet.setRowHeight(r, 20);
  sheet.getRange(r, 1, 1, NC)
    .setValues([["ÍTEM", "DESCRIPCIÓN", "UND", "CANT.", "VR. UNITARIO", "VR. TOTAL"]])
    .setBackground(NEGRO).setFontColor("#1a1a1a").setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  r++;

  // ── ÍTEMS: una sola escritura en bloque ──
  const tablaInicio = r;
  if (items.length) {
    const matriz = items.map(it => {
      const cant   = parseFloat(it.cantidad)   || 0;
      const precio = parseFloat(it.precio_apu) || 0;
      return [
        it.item_num || "",
        it.descripcion || "—",
        it.unidad || "",
        cant,
        precio,
        cant * precio,
      ];
    });
    const rango = sheet.getRange(tablaInicio, 1, matriz.length, NC);
    rango.setValues(matriz);
    rango.setFontSize(9).setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(tablaInicio, 1, matriz.length, 1).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(tablaInicio, 2, matriz.length, 1).setHorizontalAlignment("left").setWrap(true);
    sheet.getRange(tablaInicio, 3, matriz.length, 1).setHorizontalAlignment("center");
    sheet.getRange(tablaInicio, 4, matriz.length, 1).setHorizontalAlignment("right");
    sheet.getRange(tablaInicio, 5, matriz.length, 2).setHorizontalAlignment("right").setNumberFormat(MONEY);
    sheet.getRange(tablaInicio, 6, matriz.length, 1).setFontWeight("bold");
    r += matriz.length;
  } else {
    sheet.getRange(r, 1, 1, NC).merge().setValue("Sin ítems")
      .setFontSize(9).setFontColor("#888888").setHorizontalAlignment("center");
    r++;
  }

  // ── TOTALES (Subtotal + AIU + IVA + Total) ──
  const admVal  = Math.round(valorNeto * admPct  / 100);
  const impVal  = Math.round(valorNeto * impPct  / 100);
  const utilVal = Math.round(valorNeto * utilPct / 100);
  const sinIVA  = valorNeto + admVal + impVal + utilVal;
  const ivaVal  = Math.round(sinIVA * ivaPct / 100);
  const total   = sinIVA + ivaVal;

  const filasTot = [["SUBTOTAL", valorNeto]];
  if (admPct  > 0) filasTot.push(["ADMINISTRACIÓN (" + admPct  + "%)", admVal]);
  if (impPct  > 0) filasTot.push(["IMPREVISTOS ("    + impPct  + "%)", impVal]);
  if (utilPct > 0) filasTot.push(["UTILIDAD ("       + utilPct + "%)", utilVal]);
  if (ivaPct  > 0) filasTot.push(["IVA ("            + ivaPct  + "%)", ivaVal]);

  filasTot.forEach(([label, val]) => {
    sheet.setRowHeight(r, 16);
    sheet.getRange(r, 1, 1, 5).merge().setValue(label)
      .setFontSize(9).setFontColor("#333333").setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBorder(true, true, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 6).setValue(val).setNumberFormat(MONEY)
      .setFontSize(9).setFontColor("#333333").setHorizontalAlignment("right").setVerticalAlignment("middle")
      .setBorder(true, null, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  sheet.setRowHeight(r, 22);
  sheet.getRange(r, 1, 1, 5).merge().setValue("VALOR TOTAL OFERTA")
    .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(NEGRO)
    .setBorder(true, true, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(r, 6).setValue(total).setNumberFormat(MONEY)
    .setFontSize(12).setFontWeight("bold").setFontColor("#1a1a1a")
    .setHorizontalAlignment("right").setVerticalAlignment("middle").setBackground(NEGRO)
    .setBorder(true, null, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  r++;

  // ── Condiciones comerciales ──
  if (formaPago || plazoEntrega || validezOferta || noIncluye) {
    r += 1;
    sheet.setRowHeight(r, 18);
    sheet.getRange(r, 1, 1, NC).merge().setValue("CONDICIONES COMERCIALES")
      .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a").setBackground(GRIS_OSC)
      .setVerticalAlignment("middle");
    r++;
    [
      formaPago     ? ["FORMA DE PAGO",        formaPago]     : null,
      plazoEntrega  ? ["PLAZO DE ENTREGA",     plazoEntrega]  : null,
      validezOferta ? ["VALIDEZ DE LA OFERTA", validezOferta] : null,
    ].filter(Boolean).forEach(([label, val]) => {
      sheet.setRowHeight(r, 20);
      sheet.getRange(r, 1, 1, 2).merge().setValue(label)
        .setFontSize(9).setFontWeight("bold").setFontColor("#333333").setBackground("#f5f5f5").setVerticalAlignment("middle");
      sheet.getRange(r, 3, 1, 4).merge().setValue(val)
        .setFontSize(9).setFontColor("#333333").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("middle");
      r++;
    });
    if (noIncluye) {
      r++;
      sheet.getRange(r, 1, 1, NC).merge().setValue("NUESTRA OFERTA NO INCLUYE:")
        .setFontSize(9).setFontWeight("bold").setFontColor("#333333");
      r++;
      sheet.setRowHeight(r, 60);
      sheet.getRange(r, 1, 1, NC).merge().setValue(noIncluye)
        .setFontSize(9).setFontColor("#555555").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top");
      r++;
    }
  }

  // ── Firma ──
  r += 1;
  if (firmaId) {
    try {
      sheet.setRowHeight(r, 65);
      sheet.insertImage(DriveApp.getFileById(firmaId).getBlob(), 1, r).setWidth(200).setHeight(58);
    } catch(e) {}
    r++;
  }
  sheet.setRowHeight(r, 4);
  sheet.getRange(r, 1, 1, 3).merge()
    .setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  r++;
  const firmaLinea = [remitente, empresa].filter(Boolean).join("\n");
  if (firmaLinea) {
    sheet.setRowHeight(r, remitente && empresa ? 36 : 20);
    sheet.getRange(r, 1, 1, NC).merge().setValue(firmaLinea)
      .setFontSize(9).setFontWeight("bold").setVerticalAlignment("top")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
}
```

- [ ] **Step 2: Verificación manual (requiere push del usuario)**

Pídele al usuario que corra `clasp push`. Luego, en el editor:

```js
function _probarRenderCliente() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cots = listarCotizaciones();
  const cot  = getCotizacionCliente(cots[0].id);
  let tmp = ss.getSheetByName("_prueba_cliente_");
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet("_prueba_cliente_");
  llenarHojaCotizacionCliente(tmp, cot);
}
```

Expected: se crea la hoja `_prueba_cliente_` con la tabla limpia de 6 columnas, ítems ordenados por número, AIU/IVA/total al final y SIN ninguna fila de equipos/materiales/MO. Revisar visualmente y luego borrar la hoja.

- [ ] **Step 3: Verificación de escala (300 ítems)**

En el editor:

```js
function _probarEscala300() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fake = { cliente: "ESCALA", numero_oferta: "OF-TEST", fecha: "29/05/2026",
    valor_neto: 0, administracion_pct: 10, imprevistos_pct: 3, utilidad_pct: 5, iva_pct: 19, items: [] };
  for (let i = 1; i <= 300; i++) fake.items.push({ item_num: (Math.ceil(i/10)) + "." + (i%10), descripcion: "Actividad " + i, unidad: "ml", cantidad: i, precio_apu: 1000 + i });
  fake.valor_neto = fake.items.reduce((s,it)=>s+it.cantidad*it.precio_apu,0);
  let tmp = ss.getSheetByName("_escala300_"); if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet("_escala300_");
  const t0 = Date.now();
  llenarHojaCotizacionCliente(tmp, fake);
  Logger.log("300 ítems en " + (Date.now()-t0) + " ms");
}
```

Expected: termina sin error (muy por debajo del límite de 6 min) y loguea el tiempo. Borrar `_escala300_` al terminar.

- [ ] **Step 4: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: llenarHojaCotizacionCliente (tabla limpia, render en bloque)"
```

---

## Task 6: Exports con parámetro `tipo`

**Files:**
- Modify: `src/cotizacion.gs`

Las 4 funciones de export reciben `tipo` ("cliente" | "interna"). Por defecto: **cliente**. Un helper elige el renderer y la fuente de datos.

- [ ] **Step 1: Agregar helper selector y modificar `descargarPDFBase64`**

Agregar el helper al final de `src/cotizacion.gs`:

```js
// Elige datos + renderer según tipo ("cliente" por defecto).
function _prepararCotizacion(cotId, tipo) {
  if (tipo === "interna") {
    return { cot: getCotizacionCompleta(cotId), render: llenarHojaCotizacion };
  }
  return { cot: getCotizacionCliente(cotId), render: llenarHojaCotizacionCliente };
}
```

Reemplazar la cabecera de `descargarPDFBase64` (líneas ~37-47). Buscar:

```js
function descargarPDFBase64(cotId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cot = getCotizacionCompleta(cotId);
  if (!cot) return { ok: false, error: "Cotización no encontrada" };

  const tmpName = "_cot_dl_tmp_";
  let tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  llenarHojaCotizacion(tmp, cot);
```

Reemplazar por:

```js
function descargarPDFBase64(cotId, tipo) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const { cot, render } = _prepararCotizacion(cotId, tipo);
  if (!cot) return { ok: false, error: "Cotización no encontrada" };

  const tmpName = "_cot_dl_tmp_";
  let tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  render(tmp, cot);
```

- [ ] **Step 2: Modificar `exportarCotizacionPDF`**

Buscar la cabecera (líneas ~69-79):

```js
function exportarCotizacionPDF(cotId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cot = getCotizacionCompleta(cotId);
  if (!cot) return { ok: false, error: "Cotización no encontrada" };

  const tmpName = "_cot_pdf_";
  let tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  llenarHojaCotizacion(tmp, cot);
```

Reemplazar por:

```js
function exportarCotizacionPDF(cotId, tipo) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const { cot, render } = _prepararCotizacion(cotId, tipo);
  if (!cot) return { ok: false, error: "Cotización no encontrada" };

  const tmpName = "_cot_pdf_";
  let tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);

  render(tmp, cot);
```

- [ ] **Step 3: Modificar `enviarCotizacionEmail`**

Buscar (líneas ~793-804):

```js
function enviarCotizacionEmail(cotId, emailDestino) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const cot = getCotizacionCompleta(cotId);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Generar PDF
    const tmpName = "_cot_email_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    llenarHojaCotizacion(tmp, cot);
```

Reemplazar por:

```js
function enviarCotizacionEmail(cotId, emailDestino, tipo) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const { cot, render } = _prepararCotizacion(cotId, tipo);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Generar PDF
    const tmpName = "_cot_email_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    render(tmp, cot);
```

- [ ] **Step 4: Modificar `enviarYGuardarPDF`**

Buscar (líneas ~870-881):

```js
function enviarYGuardarPDF(cotId, emailDestino) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const cot = getCotizacionCompleta(cotId);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Generar PDF una sola vez
    const tmpName = "_cot_enviar_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    llenarHojaCotizacion(tmp, cot);
```

Reemplazar por:

```js
function enviarYGuardarPDF(cotId, emailDestino, tipo) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const { cot, render } = _prepararCotizacion(cotId, tipo);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Generar PDF una sola vez
    const tmpName = "_cot_enviar_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    render(tmp, cot);
```

- [ ] **Step 5: Verificación manual (requiere push del usuario)**

Pídele al usuario que corra `clasp push`. Luego, en el editor:

```js
function _probarExports() {
  const cots = listarCotizaciones();
  const id = cots[0].id;
  const cliente = descargarPDFBase64(id);            // default → cliente
  const interna = descargarPDFBase64(id, "interna"); // interna
  Logger.log("cliente ok: " + cliente.ok + " bytes: " + (cliente.base64||"").length);
  Logger.log("interna ok: " + interna.ok + " bytes: " + (interna.base64||"").length);
}
```

Expected: ambas `ok: true`. (Los tamaños difieren: la interna trae el desglose.)

- [ ] **Step 6: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: parámetro tipo (cliente|interna) en los exports de cotización"
```

---

## Task 7: UI — campos editables, línea manual y dos botones de export

**Files:**
- Modify: `src/index.html`

- [ ] **Step 1: Hacer editables `item_num` y `descripcion` en cada fila**

En `src/index.html`, reemplazar el bloque `html` de `renderFilaCot` (líneas ~2167-2177). Buscar:

```js
    let html = `<tr id="crow-${id}" class="crow-main" data-cotitem-id="${id}" style="background:#e3f2fd">
      <td style="text-align:center;font-weight:700;color:#1a237e"><span class="crow-num">${itemNum || ""}</span></td>
      <td class="mono" style="font-size:11px">${item.codigo_item || "—"}</td>
      <td style="font-weight:600">${item.descripcion || "—"}</td>
      <td style="text-align:center">${item.unidad || "—"}</td>
      <td><input type="number" class="cell-input" value="${cant}" min="1" style="width:68px"
          onchange="cambiarCantidadCot('${id}', this.value)"></td>
      <td class="num">${COP(precio)}</td>
      <td class="num vp" id="cvp-${id}" style="font-weight:700">${COP(total)}</td>
      <td><button class="btn-del" onclick="eliminarFilaCot(${id})">✕</button></td>
    </tr>`;
```

Reemplazar por (número y descripción ahora son inputs editables; descripción ocupa la columna del código):

```js
    const itemNumVal = item.item_num != null ? item.item_num : "";
    let html = `<tr id="crow-${id}" class="crow-main" data-cotitem-id="${id}" style="background:#e3f2fd">
      <td style="text-align:center">
        <input type="text" class="cell-input" value="${itemNumVal}" placeholder="1.1" style="width:48px;text-align:center;font-weight:700;color:#1a237e"
          onchange="cambiarItemNumCot('${id}', this.value)"></td>
      <td colspan="2">
        <input type="text" class="cell-input" value="${(item.descripcion || '').replace(/"/g,'&quot;')}" placeholder="Descripción..." style="width:100%;font-weight:600"
          onchange="cambiarDescripcionCot('${id}', this.value)"></td>
      <td style="text-align:center">${item.unidad || "—"}</td>
      <td><input type="number" class="cell-input" value="${cant}" min="1" style="width:68px"
          onchange="cambiarCantidadCot('${id}', this.value)"></td>
      <td class="num">${COP(precio)}</td>
      <td class="num vp" id="cvp-${id}" style="font-weight:700">${COP(total)}</td>
      <td><button class="btn-del" onclick="eliminarFilaCot(${id})">✕</button></td>
    </tr>`;
```

- [ ] **Step 2: Agregar los handlers de edición**

En `src/index.html`, justo después de la función `cambiarCantidadCot` (termina ~línea 2295), agregar:

```js
  function cambiarItemNumCot(id, val) {
    clearTimeout(window["num_" + id]);
    window["num_" + id] = setTimeout(() => {
      google.script.run.actualizarItemCotizacion(id, { item_num: val });
    }, 600);
  }

  function cambiarDescripcionCot(id, val) {
    clearTimeout(window["desc_" + id]);
    window["desc_" + id] = setTimeout(() => {
      google.script.run.actualizarItemCotizacion(id, { descripcion: val });
    }, 600);
  }
```

- [ ] **Step 3: Cambiar `cambiarCantidadCot` para usar la nueva función**

En `src/index.html`, dentro de `cambiarCantidadCot` (línea ~2293), buscar:

```js
      google.script.run.actualizarCantidadItem(id, cant);
```

Reemplazar por:

```js
      google.script.run.actualizarItemCotizacion(id, { cantidad: cant });
```

- [ ] **Step 4: Botón "+ Línea manual" y su lógica**

En `src/index.html`, en la barra donde está el botón de seleccionar APU para la cotización (cerca del botón `btn-crear-cot`/selección de APU, ~línea 993 o la barra de acciones de la cotización), agregar junto al de "Agregar APU" un botón:

```html
      <button class="btn-action" onclick="agregarLineaManualUI()">+ Línea manual</button>
```

Y agregar la función (junto a las demás de cotización, ~línea 2310):

```js
  function agregarLineaManualUI() {
    if (!cotActual) { showToast("Guarda la cotización primero.", "warning"); return; }
    showPrompt("Descripción de la línea:", "", "text", desc => {
      if (!desc) return;
      const unidad = prompt("Unidad (ml, m², und, GL):", "und") || "";
      const cant   = parseFloat(prompt("Cantidad:", "1")) || 0;
      const precio = parseFloat(prompt("Valor unitario:", "0")) || 0;
      const numero = prompt("Número de ítem (ej: 1.1):", "") || "";
      google.script.run
        .withSuccessHandler(() => {
          google.script.run.withSuccessHandler(renderCotizacion).getCotizacionCliente(cotActual.id);
        })
        .withFailureHandler(err => showToast("Error: " + err.message, "error"))
        .agregarLineaManual(cotActual.id, { item_num: numero, descripcion: desc, unidad, cantidad: cant, precio_apu: precio });
    });
  }
```

> Nota: `renderCotizacion` ya itera `cot.items` y renderiza filas; las líneas manuales (sin `equipos/materiales/...`) simplemente no muestran sub-ítems. Funciona sin cambios adicionales.

- [ ] **Step 5: Dos botones de export (cliente / interna)**

En `src/index.html`, reemplazar la barra de acciones (líneas ~1134-1141). Buscar:

```html
    <div class="actions-bar">
      <button class="btn-action btn-naranja" id="btn-pdf" onclick="descargarPDF()">
        ⬇ Descargar Cotización (.xlsx)
      </button>
      <button class="btn-action btn-celeste" id="btn-email" onclick="enviarPorCorreo()">
        ✉ Enviar y guardar Excel
      </button>
    </div>
```

Reemplazar por:

```html
    <div class="actions-bar">
      <button class="btn-action btn-naranja" id="btn-pdf" onclick="descargarPDF('cliente')">
        ⬇ Cotización para el cliente (.xlsx)
      </button>
      <button class="btn-action" id="btn-pdf-interna" onclick="descargarPDF('interna')" style="background:#607d8b;color:#fff">
        ⬇ Cotización interna / auditoría (.xlsx)
      </button>
      <button class="btn-action btn-celeste" id="btn-email" onclick="enviarPorCorreo()">
        ✉ Enviar al cliente y guardar
      </button>
    </div>
```

- [ ] **Step 6: Propagar `tipo` en `descargarPDF` y `enviarPorCorreo`**

En `src/index.html`, cambiar la firma de `descargarPDF` (línea ~2483). Buscar:

```js
  function descargarPDF() {
    if (!cotActual) { showToast("Guarda la cotización primero.", "warning"); return; }
    btnSetEstado("btn-pdf", "loading", "⏳ Generando Excel...");
```

Reemplazar por:

```js
  function descargarPDF(tipo) {
    if (!cotActual) { showToast("Guarda la cotización primero.", "warning"); return; }
    const btnId = tipo === "interna" ? "btn-pdf-interna" : "btn-pdf";
    btnSetEstado(btnId, "loading", "⏳ Generando Excel...");
```

Luego, dentro de esa misma función, cambiar las 3 referencias restantes `"btn-pdf"` por `btnId` y la llamada final. Buscar:

```js
      .withSuccessHandler(res => {
        if (!res.ok) {
          btnSetEstado("btn-pdf", "error", "✕ Error al generar");
          showToast("Error generando Excel: " + res.error, "error");
          return;
        }
```

Reemplazar `btnSetEstado("btn-pdf", ...)` por `btnSetEstado(btnId, ...)` en las tres ocurrencias dentro de la función (`"✕ Error al generar"`, `"✓ Descargado"`, `"✕ Error"`), y cambiar la llamada final:

```js
      .descargarPDFBase64(cotActual.id);
```

por:

```js
      .descargarPDFBase64(cotActual.id, tipo);
```

- [ ] **Step 7: Verificación manual (requiere push del usuario)**

Pídele al usuario que corra `clasp push` y luego abra el sidebar en la hoja:
1. Abrir una cotización existente. Editar el número de un ítem a "4.1" y su descripción; confirmar que al reabrir se conservan.
2. Agregar una "línea manual" con descripción/unidad/cantidad/valor; confirmar que aparece y suma al total.
3. Click en "Cotización para el cliente": descarga un .xlsx con tabla limpia (sin equipos/materiales/MO), ítems ordenados por número.
4. Click en "Cotización interna / auditoría": descarga el .xlsx con el desglose completo (como antes).

Expected: ambos documentos correctos; el del cliente sin desglose y ordenado.

- [ ] **Step 8: Commit**

```bash
git add src/index.html
git commit -m "feat(ui): item_num/descripcion editables, línea manual y export cliente/interna"
```

---

## Self-Review (cubierto)

- **Spec coverage:** documento cliente limpio (Task 5) · item_num/descripcion libres y editables (Task 3, Task 7) · líneas mixtas APU/manual (Task 4, Task 7) · orden numérico real (Task 1, usado en Task 2/5) · documento interno intacto (no se toca `llenarHojaCotizacion`; Task 6 solo enruta) · 300+ ítems (Task 5 batch + Task 2 lectura liviana + Task 4 escritura por lote) · exports con `tipo` (Task 6). Todos los criterios de aceptación del spec tienen tarea.
- **Sin placeholders:** todos los pasos de código traen el código completo; las verificaciones GAS traen función de prueba ejecutable.
- **Consistencia de tipos/nombres:** `compararItemNum` (Task 1) usado en `getCotizacionCliente` (Task 2) y `llenarHojaCotizacionCliente` (Task 5). `getCotizacionCliente` usado en Task 5/6/7. `actualizarItemCotizacion` (Task 3) usado en Task 7. `_prepararCotizacion` (Task 6) usa renderers de Task 5 y existente. Orden de columnas de `Cotizacion_Items` consistente (id, cotizacion_id, apu_id, item_num, descripcion, unidad, cantidad, precio_apu, valor_total) en Tasks 3/4.

## Notas de migración

- Las cotizaciones existentes tienen `item_num` numérico (1, 2, 3...). Es texto válido; `compararItemNum` los ordena igual. No requieren conversión.
- `actualizarCantidadItem` se deja en el código (compatibilidad), pero la UI pasa a usar `actualizarItemCotizacion`.
