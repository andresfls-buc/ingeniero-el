# Duplicar APU — Diseño

**Fecha:** 2026-07-16
**Rama:** dev

## Problema

El ingeniero arma muchos APUs parecidos (misma actividad, pequeños cambios). Hoy tiene
que crear cada uno desde cero. Quiere **duplicar** un APU existente y luego editar la copia.

La edición del APU (cabeza + ítems) **ya funciona** (`renderFormulario`,
`actualizarCabezaAPU`, `actualizarItemAPU`, `agregarItemAPU`, `eliminarItemAPU`). Lo único
que falta es **duplicar**.

## Requisito clave: el marcador de "copia" NO puede llegar al cliente

Al agregar un APU a una cotización, el texto que ve el cliente sale de
`apu.actividad || apu.descripcion || apu.codigo_item` (cotizacion.gs:814). El
`codigo_item` es **interno** — el documento del cliente nunca lo muestra (usa la actividad).

Por eso el marcador de copia vive **solo en `codigo_item`**:

- `descripcion`, `actividad`, `unidad`, `cliente`, `direccion` y los % de AIU se copian **limpios, sin cambios**.
- `codigo_item` de la copia = código base + sufijo `(copia)` / `(copia N)`.

Así el ingeniero distingue las copias en su lista interna, pero el cliente jamás ve "(copia)".
Sin columnas nuevas, sin migración de la hoja en producción.

## Numeración de copias

Para que muchas copias del mismo APU sean distinguibles:

- 1ª copia de `INS-01` → `INS-01 (copia)`
- 2ª → `INS-01 (copia 2)`
- 3ª → `INS-01 (copia 3)` …

Lógica:
1. Tomar el `codigo_item` de origen y **quitar** cualquier sufijo `(copia)` o `(copia N)` ya existente → código base.
2. Escanear los `codigo_item` de todos los APUs existentes.
3. Elegir el siguiente número libre: si no existe `base (copia)`, usar ese; si existe, usar `base (copia 2)`, luego `(copia 3)`, etc.

Esto también cubre "duplicar una copia": se parte del código base, no se acumula `(copia) (copia)`.

## Backend — `duplicarAPU(apuId)` en `apu.gs`

Reutiliza las funciones ya probadas y protegidas con `withLock` (las mismas que arreglaron
el bug de ids duplicados de julio 2026):

1. Leer el APU origen completo con `getAPUCompleto(apuId)`. Si no existe → error claro.
2. Calcular el `codigo_item` de la copia (base + siguiente sufijo de copia).
3. Crear la cabeza de la copia con `crearAPU(datos)`:
   - `codigo_item` = código marcado
   - `descripcion` = descripción **limpia** (código base, sin "(copia)") para que el fallback de la cotización nunca muestre el marcador
   - `unidad`, `cliente`, `direccion`, `actividad` copiados tal cual
   - `crearAPU` ya arranca subtotales en 0
4. Aplicar los % de AIU/desperdicio/HM de origen con `actualizarCabezaAPU(newId, {...})`.
5. Copiar cada ítem (`equipos`, `materiales`, `mano_obra`, `otros`) con `agregarItemAPU(newId, item)`:
   - Se pasan `tipo`, `recurso_id`, `descripcion_manual`, `cantidad`, `rendimiento`, `precio_unitario`.
   - `agregarItemAPU` recalcula `valor_parcial`, la partida y los subtotales bajo lock.
6. Devolver `{ id: newId }`.

**Por qué reutilizar `crearAPU` + `agregarItemAPU`:** heredamos el `withLock` + `flush` que
evita colisiones de id. No escribimos un copiador de filas nuevo que podría reintroducir el bug.

## Frontend — `index.html`

1. **Botón "⧉ Duplicar"** en cada fila de la lista de APUs (`renderLista`), junto a
   "Abrir →" y "✕".
2. `duplicarAPU(id)` en el cliente:
   - Muestra estado "Duplicando..." en la fila/lista.
   - Llama `google.script.run...duplicarAPU(id)`.
   - En éxito: muestra confirmación visible ("APU duplicado ✓") **y abre la copia** en el
     editor con `abrirAPU(nuevoId)`. El ingeniero nunca se pierde que pasó, y queda listo para editar.
   - En fallo: mensaje de error, se queda en la lista.

## Fuera de alcance (YAGNI)

- No se toca el flujo de edición (ya funciona).
- No se agregan columnas a la hoja APU.
- No se duplican cotizaciones (esto es solo APUs).

## Verificación

- `node --check src/apu.gs` (sintaxis).
- El usuario hace `clasp push` y prueba en el sidebar:
  1. Duplicar un APU → aparece la copia abierta, con todos los ítems y subtotales iguales al origen.
  2. En la lista, la copia se ve como `CÓDIGO (copia)`; duplicar otra vez → `(copia 2)`.
  3. Agregar la copia a una cotización cliente → el documento del cliente muestra la **actividad**, sin "(copia)".
