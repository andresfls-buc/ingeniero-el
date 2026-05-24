# Subtotales por sección en exportación de cotización

**Date:** 2026-05-24  
**Status:** Approved

## Problem

The cotización export (`llenarHojaCotizacion`) renders each APU block with sub-sections (Equipos, Materiales, Mano de Obra, Otros) and individual line items, but shows no subtotal per section — only a single "SUBTOTAL ÍT. X" row at the end of the whole APU.

The APU export (`_llenarHojaAPU`) already renders a styled subtotal row after each section. The two documents are inconsistent and the engineer cannot read section-level costs in the cotización at a glance.

## Goal

After each section's items in the cotización export, render a per-section subtotal row that matches the visual style of the APU export — gray background, bold, right-aligned, `$#,##0` currency format.

## Scope

- **In scope:** `llenarHojaCotizacion` in `src/cotizacion.gs`
- **Out of scope:** `_llenarHojaAPU` (already has subtotals, no change), data loading, column structure, AIU/total rows

## Design

### Change location

`src/cotizacion.gs` — inside the `SEC.forEach` loop in `llenarHojaCotizacion`, immediately after the `subitems.forEach(...)` block and before the existing `"SUBTOTAL ÍT. X"` row.

### Subtotal computation

```js
const secSubtotal = subitems.reduce((s, it) => s + (parseFloat(it.valor_parcial) || 0), 0);
```

Sections with zero items already `return` early, so no subtotal row is emitted for empty sections — consistent with APU behavior.

### Row style (matches APU subtotals, `apu.gs` lines 348–360)

| Element | Value |
|---|---|
| Cols 1–5 merged | Label: `"Subtotal " + sec.label` (e.g. `"Subtotal EQUIPOS"`) |
| Col 6 | `secSubtotal` formatted as `"$"#,##0` |
| Background | `GRIS_CLR` (`#f5f5f5`) |
| Font | bold, size 8, `#333333` (label) / `#111111` (value) |
| Border | `#888888 SOLID` all sides |
| Row height | 20px |

### Visual result (per APU block in cotización)

```
  EQUIPOS
    item 1 ...                              $X,XXX
    item 2 ...                              $X,XXX
                  Subtotal EQUIPOS         $XX,XXX   ← NEW
  MATERIALES
    item 1 ...                              $X,XXX
                  Subtotal MATERIALES      $XX,XXX   ← NEW
  MANO DE OBRA
    item 1 ...                              $X,XXX
                  Subtotal MANO DE OBRA    $XX,XXX   ← NEW
                  SUBTOTAL ÍT. 1          $XXX,XXX   ← existing (unchanged)
```

### Affected export surfaces

Because all three export paths call `llenarHojaCotizacion`, the fix applies to all automatically:

- XLSX download (`descargarPDFBase64`)
- PDF save to Drive (`exportarCotizacionPDF`)
- Email PDF attachment (`enviarCotizacionEmail`, `enviarYGuardarPDF`)

## Implementation

Single edit in `src/cotizacion.gs`, inside the `SEC.forEach` loop, after `subitems.forEach(...)` and before the loop ends.

Estimated diff: ~12 lines added.

No data-model changes, no new functions, no other files touched.
