# Plan de Implementación — Rediseño Cotización Cliente

> **Para el que implementa:** Ejecuta este plan tarea por tarea, en orden. Cada paso es pequeño (2-5 min). Marca cada `- [ ]` al terminarlo. NO saltes pasos. NO inventes código que no esté aquí.

**Goal:** Que la "Cotización Cliente" se arme desde UNA sola fuente de datos (las hojas `Cotizaciones` + `Cotizacion_Items` que ya existen), eliminando la hoja paralela `Cotizacion_Cliente_Items` que causa el bug `doc=null`. Se agrega numeración por capítulos (3.1, 3.2…) y el campo `objeto`.

**Architecture:** El backend (`cotizacion.gs`) ya tiene funciones limpias de fuente única: `getCotizacionCliente`, `actualizarItemCotizacion`, `agregarLineaManual`, `eliminarItemCotizacion`, `llenarHojaCotizacionCliente`, `enviarCotizacionEmail`. El trabajo es: (1) agregar 3 columnas, (2) numeración por capítulos, (3) **reconectar** la vista `v-cot-cliente` del frontend a esas funciones buenas, (4) export `.xlsx` limpio, (5) borrar lo viejo.

**Tech Stack:** Google Apps Script (`.gs`), HTML/JS sidebar (`index.html`), Google Sheets como BD. El usuario corre `clasp push` él mismo después de cada cambio.

---

## ⚠️ Cómo se verifica en este proyecto (LEER ANTES DE EMPEZAR)

No hay `pytest` ni runner automático. La verificación de cada tarea es:

1. **Tú avisas al usuario:** "Listo, hace falta `clasp push`."
2. **El usuario corre `clasp push`** (él lo hace, NUNCA lo corras tú).
3. **El usuario recarga el sidebar** y hace la acción descrita en "Verificación".
4. Se observa el resultado concreto esperado.

Para las tareas de BACKEND puro (columnas, funciones) la verificación se hace ejecutando la función desde el **editor de Apps Script** (menú ▶ Run) y mirando la hoja. Eso lo hace el usuario.

**Commits:** Después de cada tarea, haz `git add` + `git commit` de los archivos tocados. Mensajes en español, estilo de los commits existentes (`feat:`, `fix:`, `refactor:`).

---

## Mapa de archivos

| Archivo | Qué se hace |
|---------|-------------|
| `src/setup.gs` | Agregar columnas `objeto`, `capitulo_num`, `capitulo_nombre` a las definiciones + función de migración idempotente |
| `src/cotizacion.gs` | Numeración por capítulos; aceptar `objeto` en `actualizarCotizacion`; export `.xlsx` limpio desde fuente única |
| `src/index.html` | Reconectar la vista `v-cot-cliente` (líneas ~3251-3580) a las funciones de fuente única; quitar código de diagnóstico |
| `src/cotizacionCliente.gs` | DEPRECAR: borrar funciones que usan la hoja separada |
| `CLAUDE.md` | Borrar la sección "DÓNDE QUEDAMOS" del bug ya resuelto |

---

## Task 1: Agregar columnas nuevas a las definiciones de hojas

**Files:**
- Modify: `src/setup.gs:239-249` (definiciones `createSheet`)

- [ ] **Step 1: Agregar `objeto` a la hoja Cotizaciones**

En `src/setup.gs`, busca el bloque que define `Cotizaciones` (alrededor de la línea 239). Cámbialo para que la lista de columnas termine con `"objeto"`:

```javascript
  // === COTIZACIONES ===
  createSheet(ss, "Cotizaciones", [
    "id", "numero_oferta", "cliente", "direccion", "fecha",
    "valor_neto", "administracion_pct", "imprevistos_pct",
    "utilidad_pct", "iva_pct", "valor_total", "aprobada", "notas",
    "forma_pago", "plazo_entrega", "validez_oferta", "no_incluye", "objeto"
  ]);
```

- [ ] **Step 2: Agregar `capitulo_num` y `capitulo_nombre` a Cotizacion_Items**

Justo debajo, cambia la definición de `Cotizacion_Items`:

```javascript
  createSheet(ss, "Cotizacion_Items", [
    "id", "cotizacion_id", "apu_id", "item_num", "descripcion",
    "unidad", "cantidad", "precio_apu", "valor_total",
    "capitulo_num", "capitulo_nombre"
  ]);
```

- [ ] **Step 3: Commit**

```bash
git add src/setup.gs
git commit -m "feat: columnas objeto + capitulo_num/nombre en definiciones de hojas"
```

---

## Task 2: Función de migración idempotente (agrega columnas sin borrar datos)

Las hojas ya existen con datos, así que cambiar las definiciones del Task 1 NO basta (solo aplica en un setup nuevo). Hace falta una función que agregue las columnas a las hojas YA creadas.

**Files:**
- Modify: `src/setup.gs` (agregar función nueva al final del archivo)

- [ ] **Step 1: Escribir la función de migración**

Agrega esta función al final de `src/setup.gs`:

```javascript
// Migración V2: agrega objeto + capitulo_num + capitulo_nombre a hojas existentes.
// Idempotente: si la columna ya existe, no hace nada. Seguro de correr varias veces.
function migrarCotizacionV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let agregadas = 0;

  function asegurarColumna(nombreHoja, nombreColumna) {
    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf(nombreColumna) >= 0) return; // ya existe
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(nombreColumna);
    agregadas++;
  }

  asegurarColumna("Cotizaciones",     "objeto");
  asegurarColumna("Cotizacion_Items", "capitulo_num");
  asegurarColumna("Cotizacion_Items", "capitulo_nombre");

  SpreadsheetApp.getUi().alert("✅ Migración V2: " + agregadas + " columna(s) agregada(s).");
}
```

- [ ] **Step 2: Verificación (la hace el usuario)**

Avísale al usuario:
> "Necesito que: (1) hagas `clasp push`, (2) en el editor de Apps Script ejecutes la función `migrarCotizacionV2`, (3) me confirmes que la hoja `Cotizacion_Items` ahora tiene las columnas `capitulo_num` y `capitulo_nombre`, y `Cotizaciones` tiene `objeto`."

Esperado: alerta "✅ Migración V2: 3 columna(s) agregada(s)." y las columnas visibles en las hojas.

- [ ] **Step 3: Commit**

```bash
git add src/setup.gs
git commit -m "feat: migrarCotizacionV2 agrega columnas a hojas existentes (idempotente)"
```

---

## Task 3: Numeración por capítulos (renumerarCotizacion)

Hoy `agregarAPUaCotizacion` pone `item_num` como entero corrido (1, 2, 3). Lo cambiamos a numeración por capítulo: `capitulo_num.sufijo` → `3.1, 3.2, 4.1`. El sufijo lo asigna el sistema según el orden dentro del capítulo.

**Files:**
- Modify: `src/cotizacion.gs` (agregar función nueva + modificar `agregarAPUaCotizacion`)

- [ ] **Step 1: Escribir `renumerarCotizacion`**

Agrega esta función nueva en `src/cotizacion.gs` (ponla justo ANTES de `function agregarAPUaCotizacion` para que quede cerca):

```javascript
// Recalcula item_num de TODOS los ítems de una cotización como "capitulo.sufijo".
// Agrupa por capitulo_num en el orden en que aparecen las filas; dentro de cada
// grupo numera 1, 2, 3... Ej: cap "3" → 3.1, 3.2 · cap "4" → 4.1.
// Si un ítem no tiene capitulo_num, usa "1".
function renumerarCotizacion(ss, cotId) {
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const h        = data[0];
  const cIdx     = h.indexOf("cotizacion_id");
  const numIdx   = h.indexOf("item_num");
  const capIdx   = h.indexOf("capitulo_num");

  const contadorPorCapitulo = {}; // { "3": 2, "4": 1 }
  for (let i = 1; i < data.length; i++) {
    if (data[i][cIdx] != cotId) continue;
    let cap = capIdx >= 0 ? String(data[i][capIdx] || "").trim() : "";
    if (!cap) cap = "1";
    contadorPorCapitulo[cap] = (contadorPorCapitulo[cap] || 0) + 1;
    const nuevoNum = cap + "." + contadorPorCapitulo[cap];
    sheet.getRange(i + 1, numIdx + 1).setValue(nuevoNum);
  }
}
```

- [ ] **Step 2: Modificar `agregarAPUaCotizacion` para aceptar capítulo y renumerar**

En `src/cotizacion.gs`, reemplaza la firma y el cuerpo de `agregarAPUaCotizacion` (líneas ~667-712). El cambio: acepta `capituloNum` y `capituloNombre`, los guarda, y al final llama a `renumerarCotizacion` en vez de calcular `itemNum` a mano.

Reemplaza el bloque completo de la función por:

```javascript
function agregarAPUaCotizacion(cotId, apuId, cantidad, capituloNum, capituloNombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const apu = getAPUCompleto(apuId);
  if (!apu) return { ok: false };

  const sheet  = ss.getSheetByName("Cotizacion_Items");
  const data   = sheet.getDataRange().getValues();
  const lastId = data.length > 1
    ? Math.max(...data.slice(1).map(r => parseInt(r[0]) || 0))
    : 0;
  const newId  = lastId + 1;

  const cant      = parseFloat(cantidad) || 1;
  const precioAPU = parseFloat(apu.costo_neto) || 0;
  const valTotal  = cant * precioAPU;
  const cap       = String(capituloNum || "1").trim() || "1";
  const capNom    = String(capituloNombre || "").trim();

  // Orden de columnas: id, cotizacion_id, apu_id, item_num, descripcion,
  // unidad, cantidad, precio_apu, valor_total, capitulo_num, capitulo_nombre
  sheet.appendRow([
    newId, cotId, apuId, "",            // item_num se asigna en renumerarCotizacion
    apu.descripcion || apu.codigo_item || "",
    apu.unidad || "",
    cant, precioAPU, valTotal,
    cap, capNom
  ]);

  const datosCliente = heredarDatosClienteSiVacios(ss, cotId, apu);
  renumerarCotizacion(ss, cotId);
  recalcularCotizacion(ss, cotId);

  return {
    id:          newId,
    valor_total: valTotal,
    cliente:     datosCliente.cliente,
    direccion:   datosCliente.direccion,
    apu: {
      equipos:    apu.equipos    || [],
      materiales: apu.materiales || [],
      mano_obra:  apu.mano_obra  || [],
      otros:      apu.otros      || [],
    },
  };
}
```

- [ ] **Step 3: Verificación (la hace el usuario)**

Avísale:
> "Hace falta `clasp push`. Luego abre una cotización, agrega 2 APUs sin tocar capítulo. Deberían numerarse `1.1` y `1.2` (no `1` y `2`)."

Esperado: los `item_num` salen como `1.1`, `1.2`.

- [ ] **Step 4: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: numeracion por capitulos (renumerarCotizacion) en cotizacion items"
```

---

## Task 4: Asignar capítulo a un ítem existente

Permite que el usuario cambie a qué capítulo pertenece un ítem (y su nombre de capítulo), y que se renumere todo.

**Files:**
- Modify: `src/cotizacion.gs` (agregar función nueva)

- [ ] **Step 1: Escribir `asignarCapituloItem`**

Agrega en `src/cotizacion.gs`, debajo de `renumerarCotizacion`:

```javascript
// Cambia el capitulo_num y/o capitulo_nombre de un ítem y renumera la cotización.
function asignarCapituloItem(itemId, capituloNum, capituloNombre) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cotizacion_Items");
  const data  = sheet.getDataRange().getValues();
  const h      = data[0];
  const capIdx = h.indexOf("capitulo_num");
  const nomIdx = h.indexOf("capitulo_nombre");
  const cIdx   = h.indexOf("cotizacion_id");

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == itemId) {
      const cotId = data[i][cIdx];
      if (capituloNum !== undefined && capIdx >= 0) {
        sheet.getRange(i + 1, capIdx + 1).setValue(String(capituloNum || "1").trim() || "1");
      }
      if (capituloNombre !== undefined && nomIdx >= 0) {
        sheet.getRange(i + 1, nomIdx + 1).setValue(String(capituloNombre || "").trim());
      }
      renumerarCotizacion(ss, cotId);
      return { ok: true };
    }
  }
  return { ok: false };
}
```

- [ ] **Step 2: Verificación**

Esta se prueba más adelante desde la UI (Task 8). Por ahora solo confirma que el archivo no tiene errores de sintaxis: avísale al usuario que haga `clasp push` y confirme que NO sale error de compilación en Apps Script.

- [ ] **Step 3: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: asignarCapituloItem cambia capitulo y renumera"
```

---

## Task 5: Aceptar `objeto` al actualizar la cotización

`actualizarCotizacion` guarda varios campos pero NO `objeto`. Lo agregamos a la lista.

**Files:**
- Modify: `src/cotizacion.gs:795` (array `campos` dentro de `actualizarCotizacion`)

- [ ] **Step 1: Agregar `objeto` al array de campos**

En `src/cotizacion.gs`, dentro de `actualizarCotizacion`, busca esta línea (~795):

```javascript
      const campos = ["cliente","direccion","administracion_pct","imprevistos_pct","utilidad_pct","iva_pct","aprobada","notas","forma_pago","plazo_entrega","validez_oferta","no_incluye"];
```

Cámbiala por (agrega `"objeto"` al final):

```javascript
      const campos = ["cliente","direccion","administracion_pct","imprevistos_pct","utilidad_pct","iva_pct","aprobada","notas","forma_pago","plazo_entrega","validez_oferta","no_incluye","objeto"];
```

- [ ] **Step 2: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: actualizarCotizacion ahora guarda el campo objeto"
```

---

## Task 6: Export `.xlsx` limpio desde la fuente única

La vista actual llama a `exportarCotizacionClienteXlsx` (de `cotizacionCliente.gs`, que se va a borrar). Creamos un export que usa la fuente única (`getCotizacionCliente` + `llenarHojaCotizacionCliente`).

**Files:**
- Modify: `src/cotizacion.gs` (agregar función nueva al final)

- [ ] **Step 1: Escribir el nuevo export**

Agrega al final de `src/cotizacion.gs`:

```javascript
// Genera un .xlsx del documento del cliente desde la fuente única y lo guarda en Drive.
// Reusa el render limpio llenarHojaCotizacionCliente (sin desglose interno).
function exportarClienteXlsxV2(cotId) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const cot = getCotizacionCliente(cotId);
    if (!cot) return { ok: false, error: "Cotización no encontrada" };

    // Hoja temporal con el documento renderizado
    const tmpName = "_cot_xlsx_tmp_";
    let tmp = ss.getSheetByName(tmpName);
    if (tmp) ss.deleteSheet(tmp);
    tmp = ss.insertSheet(tmpName);
    llenarHojaCotizacionCliente(tmp, cot);
    SpreadsheetApp.flush();

    // Exportar SOLO esa hoja como xlsx
    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=xlsx&gid=" + tmp.getSheetId();
    const blob = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
    }).getBlob().setName("Cotizacion_" + (cot.numero_oferta || cotId) + ".xlsx");

    ss.deleteSheet(tmp);

    // Guardar en la carpeta de Drive configurada (o raíz si no hay)
    const cfg      = getConfig();
    const carpetaId = (cfg["carpeta_cotizaciones_cliente"] || "").trim();
    let archivo;
    if (carpetaId) {
      archivo = DriveApp.getFolderById(carpetaId).createFile(blob);
    } else {
      archivo = DriveApp.createFile(blob);
    }
    return { ok: true, nombre: archivo.getName(), url: archivo.getUrl() };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
```

- [ ] **Step 2: Verificación**

Se prueba desde la UI en Task 9. Por ahora: avísale al usuario que haga `clasp push` y confirme que no hay error de compilación.

- [ ] **Step 3: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: exportarClienteXlsxV2 genera xlsx desde fuente unica"
```

---

## Task 7: Reconectar apertura del documento cliente a la fuente única

Aquí se MATA el bug. La vista llama a `inicializarClienteDesdeAPUs` + `getCotizacionClienteDoc` (hoja separada, devuelve null). La cambiamos por `getCotizacionCliente` (fuente única, ya funciona).

**Files:**
- Modify: `src/index.html:3338-3436` (funciones `abrirDocCliente` y `renderDocCliente`)

- [ ] **Step 1: Reemplazar `abrirDocCliente`**

En `src/index.html`, reemplaza la función `abrirDocCliente` (líneas ~3338-3355) por esta versión que llama directo a `getCotizacionCliente`:

```javascript
function abrirDocCliente(cotId) {
  cliCotId  = cotId;
  cliFilas  = {};
  document.getElementById("v-lista").style.display       = "none";
  document.getElementById("v-cot").style.display         = "none";
  document.getElementById("v-cot-cliente").style.display = "";
  document.getElementById("cli-tbody").innerHTML         = '<tr><td colspan="7" style="padding:14px;color:#999;text-align:center">Cargando...</td></tr>';
  google.script.run
    .withSuccessHandler(renderDocCliente)
    .withFailureHandler(err => showToast("Error: " + err.message, "error"))
    .getCotizacionCliente(cotId);
}
```

- [ ] **Step 2: Reemplazar `renderDocCliente` (quitar diagnósticos + leer cfg aparte)**

`getCotizacionCliente` devuelve el objeto `cot` con un array `items` adentro (NO devuelve `cfg`). Los datos de empresa hay que pedirlos con `getConfig`. Reemplaza `renderDocCliente` (líneas ~3357-3436) por:

```javascript
function renderDocCliente(cot) {
  if (!cot) {
    document.getElementById("cli-tbody").innerHTML =
      '<tr><td colspan="7" style="padding:18px;color:#c62828;text-align:center">No se encontró la cotización.</td></tr>';
    showToast("No se encontró la cotización", "error");
    return;
  }
  cliActual = { cot: cot, items: cot.items || [] };

  // Traer datos de empresa (config) y pintar el resto cuando lleguen
  google.script.run.withSuccessHandler(cfg => {
    cfg = cfg || {};
    cliActual.cfg = cfg;

    // Membrete
    document.getElementById("cli-empresa").textContent = cfg.empresa || "";
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
    document.getElementById("cli-tbody").innerHTML = "";
    cliFilas = {};
    (cot.items || []).forEach(item => renderFilaCliente(item));
    recalcularTotalesCliente();
  }).withFailureHandler(err => showToast("Error config: " + err.message, "error"))
    .getConfig();
}
```

- [ ] **Step 3: Verificación (la hace el usuario) — ESTA ES LA PRUEBA DEL BUG**

Avísale:
> "Hace falta `clasp push` + recargar el sidebar. Entra por 'Pasar a Cotización Cliente'. Debe mostrar el documento con los ítems del APU (ya NO el mensaje 'No se encontró esta cotización' ni `doc=null`)."

Esperado: el documento renderiza con sus ítems. **El bug murió.**

- [ ] **Step 4: Commit**

```bash
git add src/index.html
git commit -m "fix: doc cliente lee de fuente unica getCotizacionCliente (mata bug doc=null)"
```

---

## Task 8: Reconectar edición de filas (campo, cantidad, precio) a la fuente única

Las funciones de edición llaman a `actualizarFilaClienteDoc` (hoja separada). Las cambiamos a `actualizarItemCotizacion` (fuente única). Además el campo de precio se llama `precio_apu` (no `precio_unitario`).

**Files:**
- Modify: `src/index.html:3438-3499` (`renderFilaCliente`, `guardarCampoCliente`, `cambiarCantCliente`, `cambiarPrecioCliente`)

- [ ] **Step 1: Arreglar `renderFilaCliente` (leer precio_apu, no precio_unitario)**

En `renderFilaCliente` (línea ~3442), busca:

```javascript
  const precio = parseFloat(item.precio_unitario) || 0;
```

Cámbiala por:

```javascript
  const precio = parseFloat(item.precio_apu) || 0;
```

- [ ] **Step 2: Reemplazar `guardarCampoCliente` para usar actualizarItemCotizacion**

Reemplaza `guardarCampoCliente` (líneas ~3476-3481) por:

```javascript
function guardarCampoCliente(id, campo, val) {
  clearTimeout(window["cli_" + id + "_" + campo]);
  window["cli_" + id + "_" + campo] = setTimeout(() => {
    google.script.run.actualizarItemCotizacion(id, { [campo]: val });
  }, 700);
}
```

- [ ] **Step 3: Arreglar el nombre de campo del precio en `cambiarPrecioCliente`**

En `cambiarPrecioCliente` (línea ~3498), busca:

```javascript
  guardarCampoCliente(id, "precio_unitario", f.precio);
```

Cámbiala por (la fuente única usa `precio_apu`):

```javascript
  guardarCampoCliente(id, "precio_apu", f.precio);
```

- [ ] **Step 4: Verificación (la hace el usuario)**

Avísale:
> "`clasp push` + recargar. En el documento cliente, cambia una cantidad y un valor unitario de una fila. Cierra y vuelve a abrir la cotización: los cambios deben persistir."

Esperado: cantidad y valor unitario editados se guardan y siguen ahí al reabrir.

- [ ] **Step 5: Commit**

```bash
git add src/index.html
git commit -m "fix: edicion de filas cliente usa actualizarItemCotizacion (precio_apu)"
```

---

## Task 9: Reconectar agregar fila, eliminar fila y export xlsx

Las últimas tres llamadas a la hoja separada: `agregarFilaClienteDoc`, `eliminarFilaClienteDoc`, `exportarCotizacionClienteXlsx`.

**Files:**
- Modify: `src/index.html:3524-3560` (`agregarFilaCliente`, `eliminarFilaCliente`, `exportarClienteXlsx`)

- [ ] **Step 1: Reemplazar `agregarFilaCliente` para usar agregarLineaManual**

Reemplaza `agregarFilaCliente` (líneas ~3524-3537) por:

```javascript
function agregarFilaCliente() {
  if (!cliCotId) return;
  google.script.run
    .withSuccessHandler(res => {
      if (!res || !res.id) return;
      const blankItem = { id: res.id, item_num: "", descripcion: "", unidad: "", cantidad: 0, precio_apu: 0 };
      renderFilaCliente(blankItem);
      recalcularTotalesCliente();
      const row = document.getElementById("cli-row-" + res.id);
      if (row) setTimeout(() => row.querySelector("input").focus(), 50);
    })
    .withFailureHandler(err => showToast("Error: " + err.message, "error"))
    .agregarLineaManual(cliCotId, { item_num: "", descripcion: "", unidad: "", cantidad: 0, precio_apu: 0 });
}
```

- [ ] **Step 2: Reemplazar la llamada de borrado en `eliminarFilaCliente`**

En `eliminarFilaCliente` (línea ~3545), busca:

```javascript
    google.script.run.eliminarFilaClienteDoc(id);
```

Cámbiala por:

```javascript
    google.script.run.eliminarItemCotizacion(id);
```

- [ ] **Step 3: Reemplazar la llamada de export en `exportarClienteXlsx`**

En `exportarClienteXlsx` (línea ~3552), busca la llamada al backend:

```javascript
    .exportarCotizacionClienteXlsx(cliCotId);
```

Cámbiala por:

```javascript
    .exportarClienteXlsxV2(cliCotId);
```

- [ ] **Step 4: Verificación (la hace el usuario)**

Avísale:
> "`clasp push` + recargar. En el documento cliente prueba: (a) '+ Agregar fila' crea una fila editable, (b) la ✕ elimina una fila, (c) '⬇ Guardar .xlsx' genera el archivo en Drive sin error."

Esperado: las tres acciones funcionan.

- [ ] **Step 5: Commit**

```bash
git add src/index.html
git commit -m "fix: agregar/eliminar/export fila cliente usan fuente unica"
```

---

## Task 10: Deprecar la hoja separada (borrar cotizacionCliente.gs)

Ya nada del frontend llama a las funciones de la hoja separada. Las borramos para que no quede código muerto que confunda.

**Files:**
- Modify: `src/cotizacionCliente.gs` (vaciar dejando solo nota)

- [ ] **Step 1: Confirmar que ninguna función vieja se usa todavía**

Run: `grep -rn "ClienteDoc\|exportarCotizacionClienteXlsx\|inicializarClienteDesdeAPUs\|getCotizacionClienteDoc" src/`

Esperado: SIN resultados en `src/index.html`. Si aparece alguno en index.html, NO continúes — vuelve a los Tasks 7-9 y arréglalo.
(Es normal que aparezcan dentro de `src/cotizacionCliente.gs`, ese archivo es el que vamos a vaciar.)

- [ ] **Step 2: Vaciar el archivo dejando una nota**

Reemplaza TODO el contenido de `src/cotizacionCliente.gs` por:

```javascript
// ─────────────────────────────────────────────────────────────────────────────
// DEPRECADO (2026-06-08): La cotización del cliente ahora usa la fuente única
// (hojas Cotizaciones + Cotizacion_Items) vía cotizacion.gs:
//   getCotizacionCliente, actualizarItemCotizacion, agregarLineaManual,
//   eliminarItemCotizacion, exportarClienteXlsxV2.
// La hoja paralela "Cotizacion_Cliente_Items" ya no se usa.
// ─────────────────────────────────────────────────────────────────────────────
```

- [ ] **Step 3: Verificación (la hace el usuario)**

Avísale:
> "`clasp push` + recargar. Abre el documento cliente otra vez para confirmar que TODO sigue funcionando (abrir, editar, agregar, eliminar, exportar) ahora que se borró el código viejo."

Esperado: todo sigue funcionando.

- [ ] **Step 4: Commit**

```bash
git add src/cotizacionCliente.gs
git commit -m "refactor: deprecar cotizacionCliente.gs (hoja separada ya no se usa)"
```

---

## Task 11: Limpiar diagnósticos temporales y CLAUDE.md

Quedan restos de diagnóstico mencionados en CLAUDE.md (la sección "DÓNDE QUEDAMOS"). Como el bug ya está resuelto, se limpian.

**Files:**
- Modify: `src/cotizacion.gs` (revertir el return con spreadsheetId de diagnóstico en `actualizarCotizacion`)
- Modify: `CLAUDE.md` (borrar la sección del bug)

- [ ] **Step 1: Revertir el return de diagnóstico en `actualizarCotizacion`**

En `src/cotizacion.gs`, dentro de `actualizarCotizacion` (~803), busca:

```javascript
      return { ok: true, spreadsheetId: ss.getId(), spreadsheetNombre: ss.getName() };
```

Cámbiala por (quita los campos de diagnóstico):

```javascript
      return { ok: true };
```

- [ ] **Step 2: Borrar la sección del bug en CLAUDE.md**

En `CLAUDE.md`, borra COMPLETA la sección que va desde `## 🔧 DÓNDE QUEDAMOS (debug en curso — 2026-06-06)` hasta el `---` que la cierra (justo antes de `## Qué es este proyecto`). El resto del archivo se conserva igual.

- [ ] **Step 3: Verificación**

Run: `grep -rn "__diag\|DÓNDE QUEDAMOS\|spreadsheetNombre" src/ CLAUDE.md`

Esperado: sin resultados (o solo comentarios inofensivos). Confirma que el sidebar sigue cargando tras `clasp push`.

- [ ] **Step 4: Commit**

```bash
git add src/cotizacion.gs CLAUDE.md
git commit -m "chore: limpiar diagnosticos temporales del bug doc cliente (resuelto)"
```

---

## Verificación final (end-to-end, la hace el usuario)

Después del Task 11, recorrido completo:

1. Crear/abrir una cotización.
2. Agregar 2-3 APUs → numeración `1.1, 1.2, 1.3`.
3. Editar una cantidad y un valor unitario → totales (AIU + IVA) se actualizan.
4. Entrar por "Pasar a Cotización Cliente" → documento renderiza con todos los ítems (sin desglose interno).
5. Descargar `.xlsx` → archivo en Drive.
6. Enviar PDF por correo (botón existente `enviarCotizacionEmail`) → llega el correo con el PDF.

Si los 6 pasos funcionan, el rediseño está completo.

---

## Notas para fases futuras (NO en este plan)

- **Selector de capítulo en la UI:** hoy el capítulo se asigna por defecto a "1". Para que el usuario elija el número de capítulo al agregar un APU (o lo cambie después con `asignarCapituloItem`), hace falta un control en la UI. Es una mejora separada.
- **Datos de empresa (logo/NIT/AIU por defecto):** se llenan en la hoja `Configuracion`. No es código.
- **Campo `objeto`:** este plan agrega la columna y su guardado (plomería), pero NO agrega aún el input para escribirlo ni su despliegue en el documento. Mostrar/editar `objeto` es una mejora separada.
