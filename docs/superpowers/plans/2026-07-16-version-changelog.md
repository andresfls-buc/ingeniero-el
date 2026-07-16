# Indicador de versión + "¿Qué hay de nuevo?" — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Mostrar la versión de la app junto al título y un modal "¿Qué hay de nuevo?" alimentado por un changelog único, con aviso "● Nuevo" que recuerda (por navegador) si el cliente ya vio la versión.

**Architecture:** Todo en `index.html`. Un arreglo `CHANGELOG` es la única fuente de verdad; `APP_VERSION = CHANGELOG[0].v`. El badge y el modal reusan el patrón de diálogo existente (`.dlg-overlay` + `display:flex/none`). La memoria del "Nuevo" usa `localStorage` con `try/catch`.

**Tech Stack:** HTML/CSS/JS vanilla dentro del sidebar/web app de Apps Script.

## Global Constraints

- **Solo frontend** (`index.html`). Cero cambios en hojas ni backend.
- Fuente única: mantener SOLO `CHANGELOG`; el número de versión sale de `CHANGELOG[0].v` (nunca hardcodear el número aparte).
- Todo acceso a `localStorage` va en `try/catch` — si falla, la versión se sigue mostrando (degradación limpia).
- El usuario hace `clasp push` él mismo; verificación runtime la hace él.

---

### Task 1: Datos del changelog + badge de versión visible

**Files:**
- Modify: `src/index.html` — CSS (junto a `.btn-open`, ~línea 52), toolbar de `v-lista` (~línea 775), bloque `<script>` (definir datos ~línea 1319), función `init` (~línea 1363).

**Interfaces:**
- Produces: `CHANGELOG` (array de `{v, fecha, cambios[]}`), `APP_VERSION` (string), `initVersion()` — pinta `vX.Y` en el badge.

- [ ] **Step 1: Agregar estilos del badge**

En `src/index.html`, después de la regla `.btn-open` (~línea 52), agregar:

```css
    .version-badge { background:none; border:none; color:rgba(255,255,255,.85); cursor:pointer; font-size:12px; display:inline-flex; align-items:center; gap:6px; padding:2px 4px; }
    .version-badge:hover { color:#fff; text-decoration:underline; }
    .version-nuevo { background:#ff5252; color:#fff; border-radius:10px; padding:1px 7px; font-size:10px; font-weight:700; }
```

- [ ] **Step 2: Agregar el badge en la barra de la vista de lista**

En `src/index.html`, reemplazar (~línea 775):

```html
    <h1>Sistema APU</h1>
```

por:

```html
    <h1>Sistema APU</h1>
    <button id="btn-version" class="version-badge" onclick="abrirNovedades()" title="¿Qué hay de nuevo?">
      <span id="version-num">v–</span>
      <span id="version-nuevo" class="version-nuevo" style="display:none">● Nuevo</span>
    </button>
```

- [ ] **Step 3: Definir CHANGELOG y APP_VERSION**

En `src/index.html`, dentro del `<script>`, justo después de la constante `COP` (~línea 1321), agregar:

```javascript
  // ── VERSIÓN / CHANGELOG (fuente única: el número sale de CHANGELOG[0].v) ──
  const CHANGELOG = [
    { v: "1.1", fecha: "16/07/2026", cambios: ["Ahora puedes duplicar APUs para reusarlos"] },
    { v: "1.0", fecha: "16/07/2026", cambios: ["Versión inicial del Sistema APU"] },
  ];
  const APP_VERSION = CHANGELOG[0].v;

  function initVersion() {
    const numEl = document.getElementById("version-num");
    if (numEl) numEl.textContent = "v" + APP_VERSION;
  }
```

- [ ] **Step 4: Llamar initVersion al arrancar**

En `src/index.html`, en la función `init` (~línea 1363), agregar la llamada:

```javascript
  function init() {
    initVersion();
    mostrarLista();
  }
```

- [ ] **Step 5: Verificación manual (la hace el usuario)**

El usuario hace `clasp push` y abre la web app: junto a "Sistema APU" aparece **"v1.1"** (sin el "Nuevo" todavía; eso llega en Task 2). El botón no rompe el layout de la barra.

- [ ] **Step 6: Commit**

```bash
git add src/index.html
git commit -m "feat: badge de version en la barra (fuente unica CHANGELOG)"
```

---

### Task 2: Modal "¿Qué hay de nuevo?" + memoria del aviso "Nuevo"

**Files:**
- Modify: `src/index.html` — HTML del modal (después de `modal-alert`, ~línea 1316), funciones JS (junto a `initVersion`).

**Interfaces:**
- Consumes: `CHANGELOG`, `APP_VERSION`, `initVersion` (Task 1).
- Produces: `abrirNovedades()`, `cerrarNovedades()`, y `initVersion()` extendida para mostrar/ocultar "● Nuevo".

- [ ] **Step 1: Agregar el modal de novedades**

En `src/index.html`, después del bloque `<!-- MODAL: ALERTA DE ÉXITO -->` (cierre en ~línea 1316), agregar:

```html
<!-- MODAL: NOVEDADES / CHANGELOG -->
<div id="modal-novedades" class="dlg-overlay">
  <div class="dlg-box">
    <div class="dlg-body">
      <p style="font-size:15px;font-weight:700;margin-bottom:12px;color:#1a237e">¿Qué hay de nuevo?</p>
      <div id="novedades-lista"></div>
    </div>
    <div class="dlg-footer" style="justify-content:center">
      <button class="dlg-btn-ok blue" onclick="cerrarNovedades()" style="min-width:110px">Entendido</button>
    </div>
  </div>
</div>
```

- [ ] **Step 2: Extender initVersion para el badge "Nuevo" y agregar las funciones del modal**

En `src/index.html`, reemplazar la función `initVersion` (de Task 1) por esta versión y agregar las funciones nuevas debajo:

```javascript
  function versionVista() {
    try { return localStorage.getItem("apu_version_vista"); } catch (e) { return null; }
  }
  function marcarVersionVista() {
    try { localStorage.setItem("apu_version_vista", APP_VERSION); } catch (e) {}
  }

  function initVersion() {
    const numEl = document.getElementById("version-num");
    if (numEl) numEl.textContent = "v" + APP_VERSION;
    const nuevoEl = document.getElementById("version-nuevo");
    if (nuevoEl) nuevoEl.style.display = (versionVista() !== APP_VERSION) ? "" : "none";
  }

  function abrirNovedades() {
    const html = CHANGELOG.map(function (e) {
      const items = e.cambios.map(function (c) { return "<li>" + c + "</li>"; }).join("");
      return '<div style="margin-bottom:14px">' +
               '<div style="font-weight:700;color:#222">v' + e.v +
                 ' <span style="font-weight:400;color:#888;font-size:12px">· ' + e.fecha + '</span></div>' +
               '<ul style="margin:6px 0 0 18px;color:#444;font-size:13px;line-height:1.5">' + items + '</ul>' +
             '</div>';
    }).join("");
    document.getElementById("novedades-lista").innerHTML = html;
    document.getElementById("modal-novedades").style.display = "flex";
    // El cliente ya vio la versión actual → recordar y ocultar el aviso.
    marcarVersionVista();
    const nuevoEl = document.getElementById("version-nuevo");
    if (nuevoEl) nuevoEl.style.display = "none";
  }

  function cerrarNovedades() {
    document.getElementById("modal-novedades").style.display = "none";
  }
```

- [ ] **Step 3: Verificación manual (la hace el usuario)**

El usuario hace `clasp push` y en la web app:
1. Aparece "v1.1 · ● Nuevo" (primera vez que ve esta función).
2. Clic en el badge → modal lista "v1.1 — Ahora puedes duplicar APUs…" y "v1.0 — Versión inicial…".
3. Cerrar y recargar → ya NO aparece "● Nuevo".
4. (Opcional) Agregar una entrada de prueba arriba del CHANGELOG y recargar → vuelve "● Nuevo"; abrirlo lo quita otra vez.

- [ ] **Step 4: Commit**

```bash
git add src/index.html
git commit -m "feat: modal ¿Que hay de nuevo? + aviso Nuevo con memoria localStorage"
```

---

## Self-Review

**Spec coverage:**
- Fuente única CHANGELOG + versión derivada → Task 1 Step 3. ✓
- Número visible junto al título → Task 1 Steps 2, 4. ✓
- Badge "● Nuevo" según versión vista → Task 2 Step 2 (`initVersion`). ✓
- Modal con lista de cambios → Task 2 Steps 1-2. ✓
- localStorage con try/catch + degradación → Task 2 Step 2 (`versionVista`/`marcarVersionVista`). ✓
- Solo frontend → ningún task toca .gs. ✓

**Placeholder scan:** sin TBD/TODO; todos los steps traen código o comando concreto. ✓

**Type consistency:** `CHANGELOG`, `APP_VERSION`, `initVersion`, `abrirNovedades`, `cerrarNovedades`, `versionVista`, `marcarVersionVista` usados con nombres idénticos entre tasks. El `onclick="abrirNovedades()"` (Task 1 Step 2) coincide con la función definida en Task 2. Nota: entre Task 1 y Task 2 el botón llama a `abrirNovedades` antes de existir; queda funcional al terminar Task 2 (ambos se despliegan juntos en el mismo `clasp push`). ✓
