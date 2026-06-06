# Cotización Cliente Document — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a new "Documento Cliente" screen — a free-form printable Colombian-style quotation the engineer fills manually, completely separate from the internal APU document.

**Architecture:** New sheet `Cotizacion_Cliente_Items` stores the client doc rows. New backend file `cotizacionCliente.gs` handles CRUD + .xlsx export. New HTML view `v-cot-cliente` renders the document with inline-editable rows, auto-calculated totals, and a print/export button. Accessed via a "📄 Doc. Cliente" button on each cotización in the list.

**Tech Stack:** Google Apps Script (.gs), Google Sheets, HTML/CSS/JS in sidebar (Google Apps Script HtmlService)

---

## File Map

| File | Action | What changes |
|------|--------|-------------|
| `src/setup.gs` | Modify | Add `migrarHojaClienteItems()` — creates `Cotizacion_Cliente_Items` sheet without resetting existing data |
| `src/cotizacionCliente.gs` | **Create** | All backend: `getCotizacionClienteDoc`, `agregarFilaClienteDoc`, `actualizarFilaClienteDoc`, `eliminarFilaClienteDoc`, `exportarCotizacionClienteXlsx` |
| `src/index.html` | Modify | Add "📄 Doc. Cliente" button in cots list; add `v-cot-cliente` view div + `<script>` block |

---

## Task 1 — Sheet Migration

**Files:**
- Modify: `src/setup.gs`

- [ ] **Add `migrarHojaClienteItems()` to `src/setup.gs`** right after the `migrarConfigEmpresa` function:

```javascript
// Crea la hoja Cotizacion_Cliente_Items si no existe.
function migrarHojaClienteItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName("Cotizacion_Cliente_Items")) {
    SpreadsheetApp.getUi().alert("✅ La hoja ya existe. No se hicieron cambios.");
    return;
  }
  const sheet = ss.insertSheet("Cotizacion_Cliente_Items");
  sheet.appendRow(["id", "cotizacion_id", "item_num", "descripcion", "unidad", "cantidad", "precio_unitario", "valor_total"]);
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert("✅ Hoja Cotizacion_Cliente_Items creada.");
}
```

- [ ] **Run it once:** Extensions → Apps Script → run `migrarHojaClienteItems()`. Verify the sheet appears in the spreadsheet.

- [ ] **Commit:**
```bash
git add src/setup.gs
git commit -m "feat: migrarHojaClienteItems crea hoja Cotizacion_Cliente_Items"
```

---

## Task 2 — Backend CRUD

**Files:**
- Create: `src/cotizacionCliente.gs`

- [ ] **Create `src/cotizacionCliente.gs`** with this complete content:

```javascript
// ─── COTIZACIÓN CLIENTE — CRUD ────────────────────────────────────────────────
// Hoja: Cotizacion_Cliente_Items
// Columnas: id, cotizacion_id, item_num, descripcion, unidad, cantidad, precio_unitario, valor_total

function _clienteSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cotizacion_Cliente_Items");
}

// Devuelve { cot: {...}, items: [...], cfg: {...} }
function getCotizacionClienteDoc(cotId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Cabecera de la cotización (cliente, fecha, numero_oferta, AIU, condiciones)
  const cotSheet = ss.getSheetByName("Cotizaciones");
  const cotData  = cotSheet.getDataRange().getValues();
  const cotH     = cotData[0];
  const cotRow   = cotData.slice(1).find(r => r[0] == cotId);
  if (!cotRow) return null;
  const cot = {};
  cotH.forEach((h, i) => cot[h] = cotRow[i]);
  ["administracion_pct", "imprevistos_pct", "utilidad_pct", "iva_pct"].forEach(k => {
    const v = parseFloat(cot[k]) || 0;
    cot[k] = (v > 0 && v < 1) ? Math.round(v * 100) : v;
  });

  // Ítems del documento cliente
  const sheet = _clienteSheet();
  const data  = sheet.getDataRange().getValues();
  let items = [];
  if (data.length > 1) {
    const h = data[0];
    items = data.slice(1)
      .filter(r => r[h.indexOf("cotizacion_id")] == cotId)
      .map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
    items.sort((a, b) => {
      // Orden numérico jerárquico: "1" < "1.1" < "1.2" < "2"
      const partsA = String(a.item_num || "").split(".").map(Number);
      const partsB = String(b.item_num || "").split(".").map(Number);
      for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
        const diff = (partsA[i] || 0) - (partsB[i] || 0);
        if (diff !== 0) return diff;
      }
      return 0;
    });
  }

  // Config de empresa (para header del documento)
  const cfg = getConfig();

  return { cot, items, cfg };
}

// Agrega una fila vacía. Devuelve { id, valor_total: 0 }
function agregarFilaClienteDoc(cotId) {
  const sheet   = _clienteSheet();
  const data    = sheet.getDataRange().getValues();
  const lastId  = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId = lastId + 1;
  sheet.appendRow([newId, cotId, "", "", "", 1, 0, 0]);
  return { id: newId, valor_total: 0 };
}

// Actualiza campos de una fila. Recalcula valor_total.
function actualizarFilaClienteDoc(itemId, cambios) {
  const sheet = _clienteSheet();
  const data  = sheet.getDataRange().getValues();
  const h     = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const row = [...data[i]];
      ["item_num", "descripcion", "unidad"].forEach(campo => {
        if (cambios[campo] !== undefined) row[h.indexOf(campo)] = cambios[campo];
      });
      const cant   = parseFloat(cambios.cantidad       ?? row[h.indexOf("cantidad")])       || 0;
      const precio = parseFloat(cambios.precio_unitario ?? row[h.indexOf("precio_unitario")]) || 0;
      row[h.indexOf("cantidad")]        = cant;
      row[h.indexOf("precio_unitario")] = precio;
      row[h.indexOf("valor_total")]     = cant * precio;
      sheet.getRange(i + 1, 1, 1, h.length).setValues([row]);
      return { ok: true, valor_total: cant * precio };
    }
  }
  return { ok: false };
}

// Elimina una fila por id.
function eliminarFilaClienteDoc(itemId) {
  const sheet = _clienteSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}
```

- [ ] **Commit:**
```bash
git add src/cotizacionCliente.gs
git commit -m "feat: CRUD Cotizacion_Cliente_Items (getCotizacionClienteDoc, agregar, actualizar, eliminar)"
```

---

## Task 3 — Export to Drive as .xlsx

**Files:**
- Modify: `src/cotizacionCliente.gs`

- [ ] **Append `exportarCotizacionClienteXlsx` to `src/cotizacionCliente.gs`:**

```javascript
// Genera un .xlsx formateado del documento cliente y lo guarda en Drive.
// Devuelve { ok, url, nombre } o { ok: false, error }
function exportarCotizacionClienteXlsx(cotId) {
  try {
    const doc = getCotizacionClienteDoc(cotId);
    if (!doc) return { ok: false, error: "Cotización no encontrada" };
    const { cot, items, cfg } = doc;

    const folderId = (cfg["carpeta_cotizaciones_cliente"] || "").trim();
    if (!folderId) return { ok: false, error: "Configura 'carpeta_cotizaciones_cliente' en la hoja Configuracion." };

    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const tmpName = "_cot_cli_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);

    // ── LLENAR HOJA ──
    const empresa = cfg["empresa"]           || "";
    const nit     = cfg["nit"]               || "";
    const tel     = cfg["telefono"]          || "";
    const email   = cfg["email_empresa"]     || "";
    const nombre  = cfg["nombre_remitente"]  || "";
    const cargo   = cfg["cargo"]             || "";
    const tp      = cfg["tarjeta_profesional"] || "";

    const fecha = cot.fecha
      ? Utilities.formatDate(new Date(cot.fecha), Session.getScriptTimeZone(), "dd 'de' MMMM 'de' yyyy")
      : "";

    let row = 1;

    // Membrete
    tmp.getRange(row, 1, 1, 6).merge().setValue(empresa).setFontSize(14).setFontWeight("bold");
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue([nit ? "NIT: " + nit : "", tel, email].filter(Boolean).join("  ·  "))
      .setFontSize(10).setFontColor("#555555");
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue("OFERTA DE PRECIOS No. " + (cot.numero_oferta || ""))
      .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("right");
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue("Bogotá, " + fecha)
      .setHorizontalAlignment("right").setFontSize(10).setFontColor("#555555");
    row += 2;

    // Destinatario
    tmp.getRange(row, 1, 1, 6).merge().setValue("Señores: " + (cot.cliente || "")).setFontWeight("bold");
    row++;
    if (cot.direccion) {
      tmp.getRange(row, 1, 1, 6).merge().setValue("Dirección: " + cot.direccion);
      row++;
    }
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue("Por medio de la presente nos permitimos presentar nuestra oferta de precios para la ejecución de las siguientes actividades:")
      .setWrap(true);
    row += 2;

    // Encabezado tabla
    const headers = ["ÍTEM", "DESCRIPCIÓN", "UND", "CANT", "VR. UNITARIO", "VR. TOTAL"];
    const hRange  = tmp.getRange(row, 1, 1, 6);
    hRange.setValues([headers])
      .setBackground("#1a237e").setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    row++;

    // Filas de ítems
    const COP = v => "$" + Math.round(parseFloat(v) || 0).toLocaleString("es-CO");
    items.forEach(item => {
      const cant  = parseFloat(item.cantidad)        || 0;
      const precio= parseFloat(item.precio_unitario) || 0;
      const total = cant * precio;
      tmp.getRange(row, 1, 1, 6).setValues([[
        item.item_num || "",
        item.descripcion || "",
        item.unidad || "",
        cant || "",
        COP(precio),
        COP(total),
      ]]);
      tmp.getRange(row, 1, 1, 2).setBorder(true, true, true, true, null, null, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
      row++;
    });
    row++;

    // Totales
    const cd   = items.reduce((s, i) => s + (parseFloat(i.cantidad) || 0) * (parseFloat(i.precio_unitario) || 0), 0);
    const admP = parseFloat(cot.administracion_pct) || 0;
    const impP = parseFloat(cot.imprevistos_pct)    || 0;
    const utlP = parseFloat(cot.utilidad_pct)       || 0;
    const ivaP = parseFloat(cot.iva_pct)            || 0;
    const adm  = cd * admP / 100;
    const imp  = cd * impP / 100;
    const utl  = cd * utlP / 100;
    const sin  = cd + adm + imp + utl;
    const iva  = sin * ivaP / 100;
    const tot  = sin + iva;

    [
      ["COSTO DIRECTO",                         COP(cd)],
      ["ADMINISTRACIÓN (" + admP + "%)",         COP(adm)],
      ["IMPREVISTOS (" + impP + "%)",            COP(imp)],
      ["UTILIDAD (" + utlP + "%)",               COP(utl)],
      ["SUBTOTAL SIN IVA",                       COP(sin)],
      ["IVA (" + ivaP + "%)",                    COP(iva)],
      ["VALOR TOTAL OFERTA",                     COP(tot)],
    ].forEach(([label, val], idx) => {
      const isTotal = idx === 6;
      tmp.getRange(row, 4, 1, 2).merge().setValue(label)
        .setHorizontalAlignment("right")
        .setFontWeight(isTotal ? "bold" : "normal")
        .setBackground(isTotal ? "#e8eaf6" : null);
      tmp.getRange(row, 6).setValue(val)
        .setHorizontalAlignment("right")
        .setFontWeight(isTotal ? "bold" : "normal")
        .setBackground(isTotal ? "#e8eaf6" : null);
      row++;
    });
    row++;

    // Condiciones comerciales
    [
      ["FORMA DE PAGO",    cot.forma_pago],
      ["PLAZO DE ENTREGA", cot.plazo_entrega],
      ["VALIDEZ OFERTA",   cot.validez_oferta],
      ["NO INCLUYE",       cot.no_incluye],
    ].filter(([, v]) => v).forEach(([label, val]) => {
      tmp.getRange(row, 1, 1, 6).merge().setValue(label + ":  " + val).setWrap(true);
      row++;
    });
    row += 2;

    // Cierre
    tmp.getRange(row, 1, 1, 6).merge().setValue("Quedamos atentos a sus consultas.  Cordialmente,");
    row += 2;
    tmp.getRange(row, 1, 1, 6).merge().setValue(nombre).setFontWeight("bold");
    row++;
    tmp.getRange(row, 1, 1, 6).merge()
      .setValue([cargo, tp ? "T.P. No. " + tp : ""].filter(Boolean).join("  ·  "));
    row++;
    tmp.getRange(row, 1, 1, 6).merge().setValue(empresa);

    // Formato columnas
    tmp.setColumnWidth(1, 55);
    tmp.setColumnWidth(2, 280);
    tmp.setColumnWidth(3, 55);
    tmp.setColumnWidth(4, 65);
    tmp.setColumnWidth(5, 110);
    tmp.setColumnWidth(6, 110);
    tmp.setRowHeights(1, tmp.getLastRow(), 22);

    SpreadsheetApp.flush();

    // Exportar como .xlsx y guardar en Drive
    const xlsxUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=xlsx&gid=" + tmp.getSheetId();
    const blob = UrlFetchApp.fetch(xlsxUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
    }).getBlob();

    const nombre_archivo = "COT-" + (cot.numero_oferta || cotId) + "_" + (cot.cliente || "Cliente").replace(/\s+/g, "_") + "_Cliente.xlsx";
    blob.setName(nombre_archivo);

    ss.deleteSheet(tmp);

    const folder = DriveApp.getFolderById(folderId);
    const file   = folder.createFile(blob);
    return { ok: true, url: file.getUrl(), nombre: nombre_archivo };

  } catch(e) {
    try {
      const bad = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_cot_cli_tmp_");
      if (bad) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(bad);
    } catch(_) {}
    return { ok: false, error: e.message };
  }
}
```

- [ ] **Commit:**
```bash
git add src/cotizacionCliente.gs
git commit -m "feat: exportarCotizacionClienteXlsx genera .xlsx y lo guarda en Drive"
```

---

## Task 4 — "📄 Doc. Cliente" Button in Cotizaciones List

**Files:**
- Modify: `src/index.html` (line ~2069 — cotizaciones list row)

- [ ] **Find this line in `src/index.html`** (inside `renderCotizaciones`):

```javascript
          <button class="btn-open" onclick="abrirCotizacion(${c.id})">Abrir →</button>
          <button class="btn-del" title="Eliminar cotización" onclick="confirmarEliminarCot(${c.id}, '${(c.numero_oferta || '').replace(/'/g,"\\'")}')">✕</button>
```

- [ ] **Replace with:**

```javascript
          <button class="btn-open" onclick="abrirCotizacion(${c.id})">Abrir →</button>
          <button class="btn-open" style="border-color:#2e7d32;color:#2e7d32" onclick="abrirDocCliente(${c.id})">📄 Cliente</button>
          <button class="btn-del" title="Eliminar cotización" onclick="confirmarEliminarCot(${c.id}, '${(c.numero_oferta || '').replace(/'/g,"\\'")}')">✕</button>
```

- [ ] **Commit:**
```bash
git add src/index.html
git commit -m "feat: botón Doc. Cliente en lista de cotizaciones"
```

---

## Task 5 — v-cot-cliente View (Document + Rows)

**Files:**
- Modify: `src/index.html` (add view div before `</body>` and script block)

- [ ] **Add the view div** just before the closing `</body>` tag in `src/index.html`:

```html
<!-- ═══════════════════════════════════════════════
     VISTA: DOCUMENTO COTIZACIÓN CLIENTE
════════════════════════════════════════════════ -->
<div id="v-cot-cliente" style="display:none">
  <div class="toolbar">
    <button class="btn btn-back" onclick="irTabCots()">← Volver</button>
    <h2 id="cli-titulo">Documento Cliente</h2>
    <div style="margin-left:auto;display:flex;gap:8px">
      <button class="btn-action btn-azul" id="btn-imprimir-cli" onclick="window.print()" style="font-size:12px;padding:7px 16px">🖨 Imprimir / PDF</button>
      <button class="btn-action btn-verde" id="btn-export-cli" onclick="exportarClienteXlsx()" style="font-size:12px;padding:7px 16px">⬇ Guardar .xlsx</button>
    </div>
  </div>

  <!-- Documento imprimible -->
  <div id="cli-doc" style="background:#fff;max-width:860px;margin:20px auto;padding:32px 40px;box-shadow:0 2px 12px rgba(0,0,0,.12);font-family:'Segoe UI',Arial,sans-serif;font-size:13px">

    <!-- Membrete -->
    <div style="border-bottom:3px solid #1a237e;padding-bottom:14px;margin-bottom:18px">
      <div id="cli-empresa" style="font-size:18px;font-weight:700;color:#1a237e"></div>
      <div id="cli-datos-empresa" style="font-size:11px;color:#555;margin-top:2px"></div>
      <div style="display:flex;justify-content:space-between;margin-top:10px">
        <div></div>
        <div style="text-align:right">
          <div id="cli-oferta" style="font-size:14px;font-weight:700;color:#1a237e"></div>
          <div id="cli-fecha" style="font-size:11px;color:#555"></div>
        </div>
      </div>
    </div>

    <!-- Destinatario -->
    <div style="margin-bottom:18px">
      <div><strong>Señores:</strong> <span id="cli-cliente"></span></div>
      <div id="cli-direccion-wrap" style="display:none"><strong>Dirección:</strong> <span id="cli-direccion"></span></div>
    </div>
    <p style="margin-bottom:18px;color:#333">Por medio de la presente nos permitimos presentar nuestra oferta de precios para la ejecución de las siguientes actividades:</p>

    <!-- Tabla de ítems -->
    <table style="width:100%;border-collapse:collapse;margin-bottom:6px" id="cli-tabla">
      <thead>
        <tr style="background:#1a237e;color:#fff;font-size:12px">
          <th style="width:52px;padding:7px 8px;text-align:center">ÍTEM</th>
          <th style="padding:7px 8px;text-align:left">DESCRIPCIÓN</th>
          <th style="width:55px;padding:7px 8px;text-align:center">UND</th>
          <th style="width:70px;padding:7px 8px;text-align:right">CANT</th>
          <th style="width:120px;padding:7px 8px;text-align:right">VR. UNITARIO</th>
          <th style="width:120px;padding:7px 8px;text-align:right">VR. TOTAL</th>
          <th class="no-print" style="width:28px"></th>
        </tr>
      </thead>
      <tbody id="cli-tbody"></tbody>
    </table>
    <button class="btn-add no-print" onclick="agregarFilaCliente()" style="margin-bottom:16px">+ Agregar fila</button>

    <!-- Totales -->
    <div style="display:flex;justify-content:flex-end;margin-bottom:20px">
      <table style="font-size:12px;border-collapse:collapse;min-width:300px">
        <tr><td style="padding:4px 12px;color:#555">COSTO DIRECTO</td><td style="padding:4px 8px;text-align:right;font-weight:600" id="cli-cd">$ 0</td></tr>
        <tr><td style="padding:4px 12px;color:#555">ADMINISTRACIÓN (<span id="cli-adm-pct-lbl">0</span>%)</td><td style="padding:4px 8px;text-align:right" id="cli-adm-val">$ 0</td></tr>
        <tr><td style="padding:4px 12px;color:#555">IMPREVISTOS (<span id="cli-imp-pct-lbl">0</span>%)</td><td style="padding:4px 8px;text-align:right" id="cli-imp-val">$ 0</td></tr>
        <tr><td style="padding:4px 12px;color:#555">UTILIDAD (<span id="cli-utl-pct-lbl">0</span>%)</td><td style="padding:4px 8px;text-align:right" id="cli-utl-val">$ 0</td></tr>
        <tr><td style="padding:4px 12px;color:#555">SUBTOTAL SIN IVA</td><td style="padding:4px 8px;text-align:right" id="cli-sin-iva">$ 0</td></tr>
        <tr><td style="padding:4px 12px;color:#555">IVA (<span id="cli-iva-pct-lbl">0</span>%)</td><td style="padding:4px 8px;text-align:right" id="cli-iva-val">$ 0</td></tr>
        <tr style="background:#e8eaf6;font-weight:700">
          <td style="padding:7px 12px">VALOR TOTAL OFERTA</td>
          <td style="padding:7px 8px;text-align:right;color:#1a237e;font-size:14px" id="cli-total">$ 0</td>
        </tr>
      </table>
    </div>

    <!-- Condiciones comerciales -->
    <div id="cli-condiciones" style="font-size:12px;border-top:1px solid #ddd;padding-top:14px;margin-bottom:20px;display:grid;grid-template-columns:160px 1fr;gap:6px 12px;color:#333">
      <!-- filled by JS -->
    </div>

    <!-- Cierre -->
    <div style="border-top:1px solid #ddd;padding-top:18px;font-size:12px;color:#333">
      <p style="margin-bottom:24px">Quedamos atentos a sus consultas. &nbsp; Cordialmente,</p>
      <div id="cli-firmante-nombre" style="font-weight:700"></div>
      <div id="cli-firmante-cargo" style="color:#555"></div>
      <div id="cli-firmante-empresa" style="color:#555"></div>
    </div>
  </div>
</div>
```

- [ ] **Add print CSS** inside the existing `<style>` block (near the top of `<head>`):

Find the line:
```css
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
```

Add right after it:
```css
    @media print {
      body > *:not(#v-cot-cliente) { display: none !important; }
      #v-cot-cliente { display: block !important; }
      #v-cot-cliente .toolbar { display: none !important; }
      .no-print { display: none !important; }
      #cli-doc { box-shadow: none !important; margin: 0 !important; padding: 16px !important; }
    }
```

- [ ] **Commit:**
```bash
git add src/index.html
git commit -m "feat: v-cot-cliente view — documento imprimible con tabla de ítems"
```

---

## Task 6 — Frontend JS (navigation, rows, totals, export)

**Files:**
- Modify: `src/index.html` (add `<script>` block just before `</body>`)

- [ ] **Add this script block** right before `</body>`:

```html
<script>
// ═══ DOCUMENTO COTIZACIÓN CLIENTE ════════════════════════════════════════════
let cliActual   = null;  // { cot, items, cfg }
let cliCotId    = null;
let cliFilas    = {};    // { id: { cant, precio, total } }

function abrirDocCliente(cotId) {
  cliCotId  = cotId;
  cliFilas  = {};
  document.getElementById("v-lista").style.display      = "none";
  document.getElementById("v-cot").style.display        = "none";
  document.getElementById("v-cot-cliente").style.display = "";
  document.getElementById("cli-tbody").innerHTML        = '<tr><td colspan="7" style="padding:14px;color:#999;text-align:center">Cargando...</td></tr>';
  google.script.run
    .withSuccessHandler(renderDocCliente)
    .withFailureHandler(err => showToast("Error: " + err.message, "error"))
    .getCotizacionClienteDoc(cotId);
}

function renderDocCliente(doc) {
  if (!doc) { showToast("No se encontró la cotización", "error"); return; }
  cliActual = doc;
  const { cot, items, cfg } = doc;

  // Membrete
  document.getElementById("cli-empresa").textContent      = cfg.empresa || "";
  document.getElementById("cli-datos-empresa").textContent =
    [cfg.nit ? "NIT: " + cfg.nit : "", cfg.telefono, cfg.email_empresa].filter(Boolean).join("  ·  ");
  document.getElementById("cli-oferta").textContent = "OFERTA DE PRECIOS No. " + (cot.numero_oferta || "");
  document.getElementById("cli-fecha").textContent  = "Bogotá, " + (cot.fecha
    ? new Date(cot.fecha).toLocaleDateString("es-CO", { day:"2-digit", month:"long", year:"numeric" })
    : "");
  document.getElementById("cli-titulo").textContent  = "Doc. Cliente — " + (cot.cliente || "");
  document.getElementById("cli-cliente").textContent = cot.cliente || "";
  const dir = cot.direccion || "";
  document.getElementById("cli-direccion-wrap").style.display = dir ? "" : "none";
  document.getElementById("cli-direccion").textContent = dir;

  // AIU labels
  document.getElementById("cli-adm-pct-lbl").textContent = parseFloat(cot.administracion_pct) || 0;
  document.getElementById("cli-imp-pct-lbl").textContent = parseFloat(cot.imprevistos_pct)    || 0;
  document.getElementById("cli-utl-pct-lbl").textContent = parseFloat(cot.utilidad_pct)       || 0;
  document.getElementById("cli-iva-pct-lbl").textContent = parseFloat(cot.iva_pct)            || 0;

  // Condiciones comerciales
  const conds = [
    ["FORMA DE PAGO",    cot.forma_pago],
    ["PLAZO DE ENTREGA", cot.plazo_entrega],
    ["VALIDEZ OFERTA",   cot.validez_oferta],
    ["NO INCLUYE",       cot.no_incluye],
  ].filter(([, v]) => v);
  document.getElementById("cli-condiciones").innerHTML = conds.map(([l, v]) =>
    `<strong>${l}:</strong><span>${v}</span>`
  ).join("") || "";

  // Firmante
  document.getElementById("cli-firmante-nombre").textContent  = cfg.nombre_remitente || "";
  document.getElementById("cli-firmante-cargo").textContent   =
    [cfg.cargo, cfg.tarjeta_profesional ? "T.P. No. " + cfg.tarjeta_profesional : ""].filter(Boolean).join("  ·  ");
  document.getElementById("cli-firmante-empresa").textContent = cfg.empresa || "";

  // Filas
  const tbody = document.getElementById("cli-tbody");
  tbody.innerHTML = "";
  cliFilas = {};
  (items || []).forEach(item => renderFilaCliente(item));
  recalcularTotalesCliente();
}

function renderFilaCliente(item) {
  const tbody  = document.getElementById("cli-tbody");
  const id     = item.id;
  const cant   = parseFloat(item.cantidad)        || 0;
  const precio = parseFloat(item.precio_unitario) || 0;
  const total  = cant * precio;
  cliFilas[id] = { cant, precio, total };

  const tr = document.createElement("tr");
  tr.id    = "cli-row-" + id;
  tr.style.borderBottom = "1px solid #eee";
  tr.innerHTML = `
    <td style="padding:5px 6px">
      <input type="text" class="cell-input" value="${item.item_num || ''}" placeholder="1.1"
        style="width:44px;text-align:center;font-weight:700;color:#1a237e"
        onchange="guardarCampoCliente(${id},'item_num',this.value)"></td>
    <td style="padding:5px 6px">
      <input type="text" class="cell-input" value="${(item.descripcion||'').replace(/"/g,'&quot;')}" placeholder="Descripción..."
        style="width:100%"
        onchange="guardarCampoCliente(${id},'descripcion',this.value)"></td>
    <td style="padding:5px 6px">
      <input type="text" class="cell-input" value="${item.unidad||''}" placeholder="und"
        style="width:44px;text-align:center"
        onchange="guardarCampoCliente(${id},'unidad',this.value)"></td>
    <td style="padding:5px 6px">
      <input type="number" class="cell-input" value="${cant||''}" min="0" placeholder="0"
        style="width:58px;text-align:right"
        onchange="cambiarCantCliente(${id},this.value)"></td>
    <td style="padding:5px 6px">
      <input type="number" class="cell-input" value="${precio||''}" min="0" placeholder="0"
        style="width:100px;text-align:right"
        onchange="cambiarPrecioCliente(${id},this.value)"></td>
    <td style="padding:5px 8px;text-align:right;font-weight:600" id="cli-vp-${id}">${COP(total)}</td>
    <td class="no-print" style="padding:3px">
      <button class="btn-del" onclick="eliminarFilaCliente(${id})">✕</button></td>`;
  tbody.appendChild(tr);
}

function guardarCampoCliente(id, campo, val) {
  clearTimeout(window["cli_" + id + "_" + campo]);
  window["cli_" + id + "_" + campo] = setTimeout(() => {
    google.script.run.actualizarFilaClienteDoc(id, { [campo]: val });
  }, 700);
}

function cambiarCantCliente(id, val) {
  const f = cliFilas[id]; if (!f) return;
  f.cant  = parseFloat(val) || 0;
  f.total = f.cant * f.precio;
  document.getElementById("cli-vp-" + id).textContent = COP(f.total);
  recalcularTotalesCliente();
  guardarCampoCliente(id, "cantidad", f.cant);
}

function cambiarPrecioCliente(id, val) {
  const f = cliFilas[id]; if (!f) return;
  f.precio = parseFloat(val) || 0;
  f.total  = f.cant * f.precio;
  document.getElementById("cli-vp-" + id).textContent = COP(f.total);
  recalcularTotalesCliente();
  guardarCampoCliente(id, "precio_unitario", f.precio);
}

function recalcularTotalesCliente() {
  if (!cliActual) return;
  const { cot } = cliActual;
  const cd   = Object.values(cliFilas).reduce((s, f) => s + (f ? f.total : 0), 0);
  const admP = parseFloat(cot.administracion_pct) || 0;
  const impP = parseFloat(cot.imprevistos_pct)    || 0;
  const utlP = parseFloat(cot.utilidad_pct)       || 0;
  const ivaP = parseFloat(cot.iva_pct)            || 0;
  const adm  = cd * admP / 100;
  const imp  = cd * impP / 100;
  const utl  = cd * utlP / 100;
  const sin  = cd + adm + imp + utl;
  const iva  = sin * ivaP / 100;
  const tot  = sin + iva;
  document.getElementById("cli-cd").textContent      = COP(cd);
  document.getElementById("cli-adm-val").textContent = COP(adm);
  document.getElementById("cli-imp-val").textContent = COP(imp);
  document.getElementById("cli-utl-val").textContent = COP(utl);
  document.getElementById("cli-sin-iva").textContent = COP(sin);
  document.getElementById("cli-iva-val").textContent = COP(iva);
  document.getElementById("cli-total").textContent   = COP(tot);
}

function agregarFilaCliente() {
  if (!cliCotId) return;
  google.script.run
    .withSuccessHandler(res => {
      if (!res || !res.id) return;
      const blankItem = { id: res.id, item_num: "", descripcion: "", unidad: "", cantidad: 0, precio_unitario: 0 };
      renderFilaCliente(blankItem);
      recalcularTotalesCliente();
      const row = document.getElementById("cli-row-" + res.id);
      if (row) setTimeout(() => row.querySelector("input").focus(), 50);
    })
    .withFailureHandler(err => showToast("Error: " + err.message, "error"))
    .agregarFilaClienteDoc(cliCotId);
}

function eliminarFilaCliente(id) {
  showConfirm("¿Eliminar esta fila?", () => {
    const row = document.getElementById("cli-row-" + id);
    if (row) row.remove();
    delete cliFilas[id];
    recalcularTotalesCliente();
    google.script.run.eliminarFilaClienteDoc(id);
  }, "Eliminar");
}

function exportarClienteXlsx() {
  if (!cliCotId) return;
  btnSetEstado("btn-export-cli", "loading", "Generando...");
  google.script.run
    .withSuccessHandler(res => {
      if (res && res.ok) {
        btnSetEstado("btn-export-cli", "success", "✓ Guardado");
        showToast("Guardado en Drive: " + res.nombre, "success");
        setTimeout(() => btnSetEstado("btn-export-cli", "", "⬇ Guardar .xlsx"), 3000);
      } else {
        btnSetEstado("btn-export-cli", "error", "✕ Error");
        showToast("Error: " + (res && res.error ? res.error : "desconocido"), "error");
      }
    })
    .withFailureHandler(err => {
      btnSetEstado("btn-export-cli", "error", "✕ Error");
      showToast("Error: " + err.message, "error");
    })
    .exportarCotizacionClienteXlsx(cliCotId);
}
</script>
```

- [ ] **Verify `irTabCots()` hides `v-cot-cliente`** — find the `irTabCots` function and add:
```javascript
document.getElementById("v-cot-cliente").style.display = "none";
```
alongside the other `style.display = "none"` lines in that function.

- [ ] **Do the same for `irTabAPUs()`** — add:
```javascript
document.getElementById("v-cot-cliente").style.display = "none";
```

- [ ] **Commit:**
```bash
git add src/index.html
git commit -m "feat: JS completo para v-cot-cliente — filas editables, totales, export"
```

---

## Task 7 — Smoke Test

- [ ] Run `clasp push`
- [ ] Open the sidebar → go to Cotizaciones tab → confirm "📄 Cliente" button appears on each row
- [ ] Click "📄 Cliente" on a cotización → confirm document loads with header (empresa, oferta No., cliente)
- [ ] Click "+ Agregar fila" → confirm blank row appears and ÍTEM field is focused
- [ ] Type an item number, description, unit, quantity, price → confirm VR. TOTAL calculates and VALOR TOTAL OFERTA updates
- [ ] Click 🖨 Imprimir → confirm only the document renders (toolbar/buttons hidden)
- [ ] Click ⬇ Guardar .xlsx → confirm file appears in the "Cotizaciones Cliente" folder in Drive
- [ ] Click ← Volver → confirm returns to cotizaciones list

---

## Self-Review

**Spec coverage check:**
- ✅ `Cotizacion_Cliente_Items` sheet — Task 1
- ✅ `getCotizacionClienteDoc` — Task 2
- ✅ `agregarFilaClienteDoc` — Task 2
- ✅ `actualizarFilaClienteDoc` — Task 2
- ✅ `eliminarFilaClienteDoc` — Task 2
- ✅ `exportarCotizacionClienteXlsx` to `carpeta_cotizaciones_cliente` — Task 3
- ✅ "📄 Doc. Cliente" button in list — Task 4
- ✅ Free-form rows, inline editable, auto-delete — Tasks 5 & 6
- ✅ Auto-calculated totals (AIU + IVA) — Task 6
- ✅ Print CSS (`@media print`) — Task 5
- ✅ Header from `Configuracion` (empresa, NIT, tel, email, nombre, cargo, TP) — Task 6
- ✅ Commercial conditions from cotización — Task 6
- ✅ `irTabCots` / `irTabAPUs` hide new view — Task 6
