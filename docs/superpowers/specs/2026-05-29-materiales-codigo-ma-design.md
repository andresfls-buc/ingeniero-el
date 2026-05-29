# Materiales mostrados como MA-xxx (consistente con EQ-/MO-/OT-)

**Fecha:** 2026-05-29
**Estado:** Diseño aprobado, pendiente implementación

## Contexto

Hoy los códigos visibles por ítem siguen una estrategia híbrida (ver
`2026-05-28-codigos-desperdicio-hm-design.md`, decisión #1):

| Tipo | Código mostrado |
|------|----------------|
| MATERIAL | `Materiales.codigo` real del proveedor (ACC1, TC15…) |
| EQUIPO | `EQ-` + id zero-padded a 3 |
| MANO_OBRA | `MO-` + id zero-padded a 3 |
| OTRO | `OT-` + id zero-padded a 3 |

El ingeniero pidió que **Materiales también use un código autogenerado `MA-xxx`**,
consistente con los demás tipos, en **todas** las vistas (BD de Precios, APU,
cotizaciones). Esto **reemplaza la decisión #1** del spec anterior: ya no se busca
preservar la trazabilidad con los códigos de proveedor en pantalla.

## Decisión

El código mostrado de un material pasa a ser `MA-` + `Materiales.id` zero-padded a
3 dígitos (id 1 → `MA-001`), paralelo a `EQ-001` / `MO-001` / `OT-001`.

El campo `codigo` seguido (ACC1, TC15…) **permanece intacto en la hoja `Materiales`**
pero ya no se muestra en ninguna parte. Sin migración de datos.

## Cambios técnicos

### `apu.gs`

1. **`resolverCodigoItem(item)`** — caso MATERIAL pasa de lookup a
   `"MA-" + _padId(id)`. El parámetro `materialesCodigoById` queda obsoleto: se
   elimina de la firma.
2. Eliminar la construcción del mapa `materialesCodigoById` (apu.gs:187-189) y su
   paso a `resolverCodigoItem` (apu.gs:339).

Esto cubre automáticamente la columna ÍTEM del **APU impreso**.

### `cotizacion.gs`

3. Eliminar la construcción del mapa `materialesCodigoById` (cotizacion.gs:154-156)
   y actualizar la llamada `resolverCodigoItem(it)` (cotizacion.gs:307).

Esto cubre la columna ÍTEM de los subítems en la **cotización impresa**.

### `index.html`

4. Añadir helper JS `padId(id)` (equivalente cliente de `_padId`).
5. Mostrar `MA-${padId(item.id)}`:
   - Tabla de materiales en BD de Precios (`~2828`).
   - Columna CÓDIGO del modal selector de materiales (`~1855`).
6. Búsqueda (`filtrarBD` `~2803`, `filtrarModal` `~1827`): seguir matcheando el
   `codigo` oculto (ACC1 aún encontrable) **y** además el nuevo `MA-xxx`.
7. Formulario de edición BD de materiales (`~2879`): el input editable `CÓDIGO`
   pasa a ser display read-only de `MA-xxx` (en registro nuevo: "se asigna al
   guardar"). `guardarRecursoBD` deja de enviar `codigo`.

### `bd.gs`

8. `_bdConf("MATERIAL").campoValorUpdate` — si `d.codigo === undefined`, devolver
   `undefined` para la columna `codigo` (preservar valor existente, no sobrescribir
   con vacío). Idéntico patrón al de MANO_OBRA/EQUIPO para campos ausentes.

## Casos borde

- **Material con `codigo` vacío en BD**: ya no importa; el display es siempre
  `MA-` + id.
- **Ítem con `recurso_id` vacío** (descripción manual): sigue mostrando `—`.
- **Registro nuevo sin id aún**: el form muestra "se asigna al guardar"; tras
  guardar, la tabla ya muestra `MA-<nuevo_id>`.

## Fuera de alcance

- Borrar la columna `codigo` de la hoja `Materiales` (se conserva).
- Cambios en desperdicio / herramienta menor / AIU.

## Criterios de éxito

1. En BD de Precios, cada material muestra `MA-001`, `MA-002`… (no ACC1).
2. El selector de materiales del APU muestra `MA-xxx`.
3. La columna ÍTEM del APU impreso y de la cotización impresa muestra `MA-xxx`
   para materiales, y `EQ-`/`MO-`/`OT-` para los demás (sin cambios).
4. Editar un material y guardarlo no borra su `codigo` original en la hoja.
