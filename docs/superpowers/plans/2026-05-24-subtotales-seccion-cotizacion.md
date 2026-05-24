# Subtotales por sección en cotización Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a per-section subtotal row (Equipos, Materiales, Mano de Obra, Otros) inside each APU block in the cotización export, matching the visual style already used in the APU export.

**Architecture:** Single edit inside `llenarHojaCotizacion` in `src/cotizacion.gs`. After the `subitems.forEach` block inside `SEC.forEach`, insert a subtotal row that sums `valor_parcial` across that section's items and renders it with the same gray/bold/border style as `_llenarHojaAPU` subtotals. No data-model changes needed — `valor_parcial` is already on each sub-item.

**Tech Stack:** Google Apps Script, Google Sheets API (SpreadsheetApp)

---

### Task 1: Add per-section subtotal rows in `llenarHojaCotizacion`

**Files:**
- Modify: `src/cotizacion.gs` — inside `SEC.forEach` loop, lines ~311–312

**Context — what the loop currently looks like (cotizacion.gs ~262–312):**

```js
SEC.forEach(sec => {
  const subitems = item[sec.key] || [];
  if (!subitems.length) return;

  // section header row ...
  r++;

  // item rows
  subitems.forEach((it, i) => {
    // ... render each line item ...
    r++;
  });
  // ← INSERT SUBTOTAL ROW HERE (before the closing }); of SEC.forEach)
});
```

- [ ] **Step 1: Insert the subtotal row block**

In `src/cotizacion.gs`, find the closing `});` of `subitems.forEach` (currently line 311) and add the following block immediately after it, before the `});` that closes `SEC.forEach` (currently line 312):

```js
      // Subtotal de sección
      const secSubtotal = subitems.reduce((s, it) => s + (parseFloat(it.valor_parcial) || 0), 0);
      sheet.setRowHeight(r, 20);
      sheet.getRange(r, 1, 1, 5).merge()
        .setValue("Subtotal " + sec.label)
        .setFontSize(8).setFontWeight("bold").setFontColor("#333333")
        .setHorizontalAlignment("right").setVerticalAlignment("middle")
        .setBackground(GRIS_CLR)
        .setBorder(true, true, true, null, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(r, 6).setValue(secSubtotal).setNumberFormat(MONEY)
        .setFontSize(8).setFontWeight("bold").setFontColor("#111111")
        .setHorizontalAlignment("right").setVerticalAlignment("middle")
        .setBackground(GRIS_CLR)
        .setBorder(true, null, true, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
      r++;
```

After the edit the loop should look like this:

```js
      subitems.forEach((it, i) => {
        // ... existing item rendering ...
        r++;
      });

      // Subtotal de sección
      const secSubtotal = subitems.reduce((s, it) => s + (parseFloat(it.valor_parcial) || 0), 0);
      sheet.setRowHeight(r, 20);
      sheet.getRange(r, 1, 1, 5).merge()
        .setValue("Subtotal " + sec.label)
        .setFontSize(8).setFontWeight("bold").setFontColor("#333333")
        .setHorizontalAlignment("right").setVerticalAlignment("middle")
        .setBackground(GRIS_CLR)
        .setBorder(true, true, true, null, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(r, 6).setValue(secSubtotal).setNumberFormat(MONEY)
        .setFontSize(8).setFontWeight("bold").setFontColor("#111111")
        .setHorizontalAlignment("right").setVerticalAlignment("middle")
        .setBackground(GRIS_CLR)
        .setBorder(true, null, true, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID);
      r++;
    });
```

- [ ] **Step 2: Push to Google Apps Script**

Run in terminal (user executes this themselves):
```
clasp push
```

Expected: `Pushed N files.` with no errors.

- [ ] **Step 3: Verify in the webapp**

Open the webapp → go to Cotizaciones → open any cotización that has items in at least 2 sections (e.g. Materiales + Mano de Obra) → click Descargar XLSX or Exportar PDF.

Expected result in the downloaded file:
- After each section's line items, a gray row labeled e.g. `"Subtotal MATERIALES"` with the section total in col F.
- Empty sections have no subtotal row (they are already skipped by `if (!subitems.length) return`).
- The existing `"SUBTOTAL ÍT. X"` row (whole-APU total) is still present and unchanged below all sections.
- Visual style matches the APU export subtotals: `#f5f5f5` background, bold, `#888888` border.

- [ ] **Step 4: Commit**

```bash
git add src/cotizacion.gs
git commit -m "feat: add per-section subtotals in cotización export

Adds a subtotal row (Equipos / Materiales / Mano de Obra / Otros)
after each section's items inside every APU block of the cotización
export. Style matches the APU export subtotals. Applies to XLSX
download, PDF save to Drive, and email attachment.

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
```

---

### Task 2: (Already done) Fix cropped descriptions in exports

> This task was completed before the plan was written.

**Files modified:** `src/apu.gs`, `src/cotizacion.gs`

**What was done:** Removed `sheet.setRowHeight(r, 20)` for item rows in `_llenarHojaAPU` and `sheet.setRowHeight(r, 17)` for sub-item rows in `llenarHojaCotizacion`. With `setWrap(true)` already in place, removing the fixed height allows Google Sheets to auto-size rows whose descriptions span multiple lines.

- [x] Fix applied and tested via code review.
- [ ] **Verify with clasp push** — after pushing Task 1, also confirm that long descriptions (e.g. "Gases refrigerantes, presurización, soldadura 5%y aceite") now show fully without cropping.
