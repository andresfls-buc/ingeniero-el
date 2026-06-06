# Spec: Documento de Cotización para el Cliente

**Date:** 2026-06-06
**Status:** Approved

---

## Context

The engineer currently has an internal cotización screen (with full APU breakdown) and an internal export (.xlsx). The client must never see the APU breakdown (costs of mano de obra, materiales, equipos, etc. — that is proprietary).

What's needed: a **separate, free-form document** the engineer fills manually to match whatever structure the client's licitación asks for. Each row is typed by the engineer — description, unit, quantity, unit price. The system auto-calculates totals (AIU + IVA) and renders a clean professional Colombian-style quotation. The document is saved as `.xlsx` to a dedicated Google Drive folder.

---

## What We Are NOT Building

- Auto-population from APU items into the client document (the engineer decides what goes in each row)
- A modal or popup for adding rows (inline editing only)
- A separate "company info" form per document (filled once in Configuración, reused everywhere)

---

## 1. Configuración — New Fields

Add to the existing `Configuración` sheet (and to the "Mi Empresa" UI section):

| clave | descripción |
|-------|-------------|
| `nit` | NIT de la empresa (ej: 900.123.456-7) |
| `telefono` | Teléfono empresa |
| `email_empresa` | Email de contacto |
| `cargo` | Cargo del firmante (ej: Ingeniero Mecánico) |
| `tarjeta_profesional` | No. Tarjeta Profesional |
| `carpeta_cotizaciones_cliente` | ID carpeta Drive para docs cliente |

The existing `carpeta_cotizaciones` becomes the internal cotizaciones folder.
The existing `empresa`, `nombre_remitente`, `logo_id` are already there.

**UI:** Add a "⚙ Mi Empresa" section/tab in the sidebar where ALL Configuración fields are editable in one place. Fill once, flows to every document.

---

## 2. New Sheet: `Cotizacion_Cliente_Items`

Columns: `id, cotizacion_id, item_num, descripcion, unidad, cantidad, precio_unitario, valor_total`

- Completely separate from `Cotizacion_Items` (which holds internal APU-linked rows)
- `valor_total = cantidad × precio_unitario` (calculated on write)
- No `apu_id` — these rows are always free-form

---

## 3. New Backend Functions (cotizacion.gs)

```
getCotizacionClienteDoc(cotId)
  → returns: cot header fields + Cotizacion_Cliente_Items rows + Configuración data

agregarFilaClienteDoc(cotId, datos)
  → inserts row into Cotizacion_Cliente_Items, returns { id, valor_total }

actualizarFilaClienteDoc(itemId, cambios)
  → updates item_num, descripcion, unidad, cantidad, precio_unitario, valor_total

eliminarFilaClienteDoc(itemId)
  → deletes row from Cotizacion_Cliente_Items

exportarCotizacionClienteXlsx(cotId)
  → builds formatted .xlsx, saves to carpeta_cotizaciones_cliente in Drive
  → returns { url, nombre } of the saved file
```

The AIU percentages used in the client doc are read from the `Cotizaciones` row (same as internal) — they are displayed but the engineer can override them per session (not saved separately; if they need to be different they edit the cotización first).

---

## 4. New View: `v-cot-cliente`

### Access
From the cotizaciones list, each row gets a second button: **"📄 Doc. Cliente"** (alongside the existing "Abrir" button). Clicking it navigates to `v-cot-cliente`.

### Layout

```
TOOLBAR
  ← Volver  |  "Doc. Cliente — [nombre cliente]"  |  [🖨 Imprimir]  [⬇ Descargar .xlsx]

DOCUMENT BODY (white card, max-width 900px, centered, looks like paper)
┌─────────────────────────────────────────────────────────────┐
│  [LOGO]   EMPRESA S.A.S · NIT 900.XXX.XXX-X                 │
│           Bogotá · Tel 301-XXX · email@empresa.com           │
│                                                             │
│  OFERTA DE PRECIOS No. 2026-001                             │
│  Bogotá, 06 de junio de 2026                               │
├─────────────────────────────────────────────────────────────┤
│  Señores: CLIENTE S.A.                                      │
│  Dirección: Cra. XX # XX-XX, Bogotá                        │
│                                                             │
│  Por medio de la presente nos permitimos presentar nuestra  │
│  oferta de precios para la ejecución de las siguientes      │
│  actividades:                                               │
├──────┬──────────────────────────┬─────┬──────┬────────┬─────┤
│ ÍTEM │ DESCRIPCIÓN              │ UND │ CANT │VR.UNIT │TOTAL│
├──────┼──────────────────────────┼─────┼──────┼────────┼─────┤
│[inp] │ [input text            ] │[inp]│[inp] │ [inp]  │ $   │
│[inp] │ [input text            ] │[inp]│[inp] │ [inp]  │ $   │
│              [+ Agregar fila]                               │
├─────────────────────────────────────────────────────────────┤
│                       COSTO DIRECTO:        $ xxx           │
│                       ADMINISTRACIÓN (10%): $ xxx           │
│                       IMPREVISTOS (5%):     $ xxx           │
│                       UTILIDAD (12%):       $ xxx           │
│                       SUBTOTAL SIN IVA:     $ xxx           │
│                       IVA (19%):            $ xxx           │
│                  ┌───────────────────────────────┐          │
│                  │  VALOR TOTAL OFERTA: $ xxx    │          │
│                  └───────────────────────────────┘          │
├─────────────────────────────────────────────────────────────┤
│  FORMA DE PAGO:     [editable]                              │
│  PLAZO DE ENTREGA:  [editable]                              │
│  VALIDEZ OFERTA:    [editable]                              │
│  NO INCLUYE:        [editable]                              │
├─────────────────────────────────────────────────────────────┤
│  Quedamos atentos a sus consultas.  Cordialmente,           │
│                                                             │
│  [Nombre Ingeniero]                                         │
│  [Cargo] · T.P. No. [tarjeta_profesional]                  │
│  [Empresa]                                                  │
└─────────────────────────────────────────────────────────────┘
```

### Row behavior
- Every field in each row is an `<input>` (inline, no modal)
- VR. TOTAL = CANT × VR. UNITARIO, recalculated instantly on any change
- Auto-save to backend (debounced 800ms) on every field change
- Delete button (✕) on each row, with confirm dialog
- "+ Agregar fila" calls `agregarFilaClienteDoc` → on success inserts row with real ID, focuses ÍTEM field

### Totals
- AIU % shown as editable inputs (seeded from cotización, changes stay local to this session — not persisted separately)
- If the engineer needs different AIU than the internal cotización, they edit the cotización first
- All totals recalculated in JS on any change

### Print / Export
- **🖨 Imprimir:** calls `window.print()`. CSS `@media print` hides toolbar, sidebar chrome; shows only the document body as a clean A4 page.
- **⬇ Descargar .xlsx:** calls `exportarCotizacionClienteXlsx(cotId)`, shows loading state, on success shows toast with link to Drive file.

---

## 5. Drive Folder Structure

```
📁 carpeta_apus               — APU exports (existing)
📁 carpeta_cotizaciones       — Cotizaciones internas exports (existing)
📁 carpeta_cotizaciones_cliente — Cotizaciones cliente exports (NEW)
```

File naming: `COT-[numero_oferta]_[cliente]_Cliente.xlsx`
Example: `COT-2026-001_Empresa_SA_Cliente.xlsx`

---

## 6. "Mi Empresa" Settings Screen

A new tab or section in the sidebar (`v-config`) where the engineer fills in ALL Configuración fields once:

- Empresa / Razón social
- NIT
- Teléfono
- Email
- Nombre del firmante (nombre_remitente)
- Cargo
- No. Tarjeta Profesional
- Logo (Drive file ID or upload)
- Carpeta APUs (Drive folder ID)
- Carpeta Cotizaciones Internas (Drive folder ID)
- Carpeta Cotizaciones Cliente (Drive folder ID)

Save button writes all fields to the `Configuración` sheet at once.

---

## 7. Scope / Out of Scope

**In scope:**
- `v-cot-cliente` view (new HTML section)
- `Cotizacion_Cliente_Items` sheet + CRUD backend functions
- `exportarCotizacionClienteXlsx` for client docs
- Configuración new fields + "Mi Empresa" UI
- `@media print` CSS for clean A4 print

**Out of scope for this sprint:**
- Updating the existing internal cotización export (it already works)
- Email sending
- Signature image upload
- PDF generation server-side (print-to-PDF via browser is sufficient)
