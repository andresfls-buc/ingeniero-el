# Cotización PDF — Una Sola Página Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Hacer que el PDF exportado de una cotización siempre quepa en exactamente 1 página, usando un layout compacto + el parámetro `scale=4` de Google Sheets.

**Architecture:** Dos cambios en `src/cotizacion.gs`: (1) reducir alturas de fila, fuentes y tamaño del logo en `llenarHojaCotizacion`; (2) reemplazar `fitw=true` por `scale=4` y bajar márgenes a 0.50" en las URLs de exportación de `enviarCotizacionEmail` y `enviarYGuardarPDF`. Sin cambios de lógica de negocio ni de estructura de datos.

**Tech Stack:** Google Apps Script (.gs), Google Sheets export API

---

## Contexto del proyecto

- El proyecto es Google Apps Script — no hay framework de tests. La verificación es manual: el usuario hace `clasp push` y prueba en el Google Sheet.
- **Nunca ejecutes `clasp push` tú mismo** — el usuario lo hace él.
- Todos los cambios van en `src/cotizacion.gs`.

---

### Task 1: Actualizar URLs de exportación PDF

**Files:**
- Modify: `src/cotizacion.gs:726-730` (función `enviarCotizacionEmail`)
- Modify: `src/cotizacion.gs:803-807` (función `enviarYGuardarPDF`)

- [ ] **Step 1: Cambiar URL en `enviarCotizacionEmail`**

Buscar el bloque en línea ~726 dentro de la función `enviarCotizacionEmail`:

```javascript
    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=pdf&gid=" + tmp.getSheetId()
      + "&portrait=true&fitw=true&size=letter"
      + "&gridlines=false&printtitle=false&sheetnames=false"
      + "&top_margin=0.75&bottom_margin=0.75&left_margin=0.75&right_margin=0.75";
```

Reemplazar por:

```javascript
    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=pdf&gid=" + tmp.getSheetId()
      + "&portrait=true&scale=4&size=letter"
      + "&gridlines=false&printtitle=false&sheetnames=false"
      + "&top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50";
```

- [ ] **Step 2: Cambiar URL en `enviarYGuardarPDF`**

Buscar el bloque en línea ~803 dentro de la función `enviarYGuardarPDF`:

```javascript
    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=pdf&gid=" + tmp.getSheetId()
      + "&portrait=true&fitw=true&size=letter"
      + "&gridlines=false&printtitle=false&sheetnames=false"
      + "&top_margin=0.75&bottom_margin=0.75&left_margin=0.75&right_margin=0.75";
```

Reemplazar por:

```javascript
    const exportUrl = "https://docs.google.com/spreadsheets/d/" + ss.getId()
      + "/export?format=pdf&gid=" + tmp.getSheetId()
      + "&portrait=true&scale=4&size=letter"
      + "&gridlines=false&printtitle=false&sheetnames=false"
      + "&top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50";
```

- [ ] **Step 3: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: use scale=4 and smaller margins for single-page PDF export"
```

---

### Task 2: Compactar layout en `llenarHojaCotizacion`

**Files:**
- Modify: `src/cotizacion.gs:126-464` (función `llenarHojaCotizacion` completa)

#### 2a — Fila logo y encabezado

- [ ] **Step 1: Reducir altura del logo y tamaño de imagen**

Buscar (~línea 162):
```javascript
  sheet.setRowHeight(r, 85);
```
Reemplazar por:
```javascript
  sheet.setRowHeight(r, 55);
```

Buscar (~línea 169):
```javascript
      img.setWidth(215).setHeight(78);
```
Reemplazar por:
```javascript
      img.setWidth(160).setHeight(48);
```

#### 2b — Fila cliente/dirección, número de oferta

- [ ] **Step 2: Reducir fila cliente**

Buscar (~línea 179):
```javascript
  sheet.setRowHeight(r, 34);
```
Reemplazar por:
```javascript
  sheet.setRowHeight(r, 22);
```

- [ ] **Step 3: Reducir fila número de oferta**

Buscar (~línea 192):
```javascript
  sheet.setRowHeight(r, 20);
```
Reemplazar por (hay dos ocurrencias antes del encabezado de tabla — esta es la primera, dentro del bloque "Datos de oferta"):
```javascript
  sheet.setRowHeight(r, 14);
```

#### 2c — Encabezado de tabla

- [ ] **Step 4: Compactar encabezado de tabla**

Buscar (~línea 210):
```javascript
  sheet.setRowHeight(r, 26);
  sheet.getRange(r, 1, 1, NC)
    .setValues([["ÍTEM", "DESCRIPCIÓN", "UND", "CANT.", "VR. UNITARIO", "VR. TOTAL"]])
    .setBackground(NEGRO).setFontColor("#1a1a1a").setFontWeight("bold").setFontSize(9)
```
Reemplazar por:
```javascript
  sheet.setRowHeight(r, 20);
  sheet.getRange(r, 1, 1, NC)
    .setValues([["ÍTEM", "DESCRIPCIÓN", "UND", "CANT.", "VR. UNITARIO", "VR. TOTAL"]])
    .setBackground(NEGRO).setFontColor("#1a1a1a").setFontWeight("bold").setFontSize(8)
```

#### 2d — Fila principal de cada APU

- [ ] **Step 5: Reducir fila principal del APU**

Buscar (~línea 241):
```javascript
    sheet.setRowHeight(r, 24);
    sheet.getRange(r, 1, 1, NC)
      .setBackground(GRIS_OSC)
```
Reemplazar por:
```javascript
    sheet.setRowHeight(r, 18);
    sheet.getRange(r, 1, 1, NC)
      .setBackground(GRIS_OSC)
```

#### 2e — Encabezados de sección (EQUIPOS, MATERIALES, etc.)

- [ ] **Step 6: Reducir encabezado de sección**

Buscar (~línea 267):
```javascript
      sheet.setRowHeight(r, 17);
```
Reemplazar por:
```javascript
      sheet.setRowHeight(r, 12);
```

#### 2f — Sub-ítems (altura explícita)

- [ ] **Step 7: Fijar altura de filas de sub-ítems**

Dentro del bloque `subitems.forEach((it, i) => { ... r++; })`, añadir `sheet.setRowHeight(r, 13);` justo antes del `r++` al final del forEach. El bloque completo del forEach queda así:

```javascript
      subitems.forEach((it, i) => {
        const bg = i % 2 === 0 ? "#ffffff" : "#fafafa";
        const cantTexto = it.cantidad != null ? it.cantidad : "";
        const mostrarUnidad = sec.key === "materiales";
        sheet.getRange(r, 1).setValue("").setBackground(bg)
          .setBorder(true, true, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        if (mostrarUnidad) {
          sheet.getRange(r, 2).setValue("      " + (it.descripcion_manual || "—"))
            .setFontSize(8).setFontColor("#222222").setHorizontalAlignment("left")
            .setVerticalAlignment("middle").setWrap(true).setBackground(bg)
            .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
          sheet.getRange(r, 3).setValue(it.unidad || "")
            .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("center")
            .setVerticalAlignment("middle").setBackground(bg)
            .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        } else {
          sheet.getRange(r, 2, 1, 2).merge().setValue("      " + (it.descripcion_manual || "—"))
            .setFontSize(8).setFontColor("#222222").setHorizontalAlignment("left")
            .setVerticalAlignment("middle").setWrap(true).setBackground(bg)
            .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        }
        sheet.getRange(r, 4).setValue(cantTexto)
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 5).setValue(parseFloat(it.precio_unitario) || 0).setNumberFormat(MONEY)
          .setFontSize(8).setFontColor("#444444").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, null, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(r, 6).setValue(parseFloat(it.valor_parcial) || 0).setNumberFormat(MONEY)
          .setFontSize(8).setFontWeight("bold").setFontColor("#111111").setHorizontalAlignment("right")
          .setVerticalAlignment("middle").setBackground(bg)
          .setBorder(true, null, true, true, null, null, BORDE, SpreadsheetApp.BorderStyle.SOLID);
        sheet.setRowHeight(r, 13);
        r++;
      });
```

#### 2g — Subtotales de sección y de APU

- [ ] **Step 8: Reducir subtotal de sección**

Buscar el bloque de subtotal de sección (~línea 315):
```javascript
      sheet.setRowHeight(r, 20);
      sheet.getRange(r, 1, 1, 5).merge()
        .setValue("Subtotal " + sec.label)
```
Reemplazar por:
```javascript
      sheet.setRowHeight(r, 15);
      sheet.getRange(r, 1, 1, 5).merge()
        .setValue("Subtotal " + sec.label)
```

- [ ] **Step 9: Reducir subtotal de APU**

Buscar el bloque de subtotal del APU (~línea 331):
```javascript
    sheet.setRowHeight(r, 20);
    sheet.getRange(r, 1, 1, 5).merge()
      .setValue("SUBTOTAL ÍT. " + (idx + 1))
```
Reemplazar por:
```javascript
    sheet.setRowHeight(r, 15);
    sheet.getRange(r, 1, 1, 5).merge()
      .setValue("SUBTOTAL ÍT. " + (idx + 1))
```

#### 2h — TOTAL COSTOS DIRECTOS, AIU, VALOR TOTAL

- [ ] **Step 10: Reducir TOTAL COSTOS DIRECTOS**

Buscar (~línea 347):
```javascript
  sheet.setRowHeight(r, 24);
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("TOTAL COSTOS DIRECTOS")
    .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
```
Reemplazar por:
```javascript
  sheet.setRowHeight(r, 18);
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("TOTAL COSTOS DIRECTOS")
    .setFontSize(9).setFontWeight("bold").setFontColor("#1a1a1a")
```

- [ ] **Step 11: Reducir filas de AIU**

Buscar (~línea 375) dentro del `aiuFilas.forEach`:
```javascript
    sheet.setRowHeight(r, 20);
```
Reemplazar por:
```javascript
    sheet.setRowHeight(r, 15);
```

- [ ] **Step 12: Reducir VALOR TOTAL OFERTA**

Buscar (~línea 388):
```javascript
  sheet.setRowHeight(r, 28);
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("VALOR TOTAL OFERTA")
    .setFontSize(12).setFontWeight("bold").setFontColor("#1a1a1a")
```
Reemplazar por:
```javascript
  sheet.setRowHeight(r, 20);
  sheet.getRange(r, 1, 1, 5).merge()
    .setValue("VALOR TOTAL OFERTA")
    .setFontSize(10).setFontWeight("bold").setFontColor("#1a1a1a")
```

#### 2i — Espaciadores

- [ ] **Step 13: Reducir espaciador antes de CONDICIONES COMERCIALES**

Buscar (~línea 401) dentro del bloque `if (formaPago || plazoEntrega || validezOferta || noIncluye)`:
```javascript
    r += 2;
    sheet.setRowHeight(r, 22);
    sheet.getRange(r, 1, 1, NC).merge()
      .setValue("CONDICIONES COMERCIALES")
```
Reemplazar por:
```javascript
    r += 1;
    sheet.setRowHeight(r, 18);
    sheet.getRange(r, 1, 1, NC).merge()
      .setValue("CONDICIONES COMERCIALES")
```

- [ ] **Step 14: Reducir espaciador antes de FIRMA**

Buscar (~línea 441):
```javascript
  r += 2;
  const firmaId = (cfg["firma_id"] || "").trim();
```
Reemplazar por:
```javascript
  r += 1;
  const firmaId = (cfg["firma_id"] || "").trim();
```

- [ ] **Step 15: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: compact llenarHojaCotizacion layout for single-page PDF"
```

---

### Task 3: Verificación manual

- [ ] **Step 1: Hacer clasp push** (el usuario lo ejecuta, no el agente)

Decirle al usuario que corra `clasp push` desde su terminal.

- [ ] **Step 2: Abrir el Google Sheet y exportar una cotización con 6+ ítems**

En la webapp, buscar una cotización con varios APUs → botón "Enviar por email" o "Guardar PDF".

- [ ] **Step 3: Confirmar que el PDF tiene exactamente 1 página**

Abrir el PDF en Drive o el adjunto del email y verificar que todo el contenido está en 1 página carta.

- [ ] **Step 4: Si el PDF tiene 2 páginas** (cotización muy densa)

El `scale=4` ya está activo, así que el problema sería que Google Sheets no está aplicando el parámetro. Verificar que la URL no tenga `fitw=true` (los dos debían haberse removido en Task 1). Si el texto es demasiado pequeño para leer, considerar cambiar a orientación horizontal añadiendo `&portrait=false` en las URLs.
