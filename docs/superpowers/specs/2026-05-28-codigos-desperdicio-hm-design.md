# Códigos visibles, Factor de Desperdicio y Herramienta Menor en APUs

**Fecha:** 2026-05-28
**Estado:** Diseño aprobado, pendiente plan de implementación

## Contexto

El cliente compartió un APU de otro software de cotización. Tres prácticas suyas mejoran la legibilidad y precisión de nuestros APUs:

1. **Códigos visibles por ítem** (`MAT-001`, `MO-001`, `EQ-001`) — conecta cada renglón con la BD de recursos sin tener que leer la descripción completa.
2. **Factor de desperdicio** declarado en el encabezado de Materiales — transparencia sobre el ajuste aplicado.
3. **Herramienta Menor (HM)** como línea automática = % del subtotal de Mano de Obra — estándar de la industria de construcción en Colombia.

Hoy el sistema no muestra códigos (columna ÍTEM es solo 1, 2, 3…), no tiene desperdicio, y la HM se debe meter a mano como "Otro".

## Decisiones de diseño

### 1. Códigos visibles — estrategia híbrida

| Tipo | Código mostrado | Fuente |
|------|----------------|--------|
| MATERIAL | `Materiales.codigo` (lookup por `recurso_id`) | BD existente (AA1, TC15, VAL-32-1A) |
| EQUIPO | `EQ-` + id zero-padded a 3 | Autogenerado desde `Equipos.id` |
| MANO_OBRA | `MO-` + id zero-padded a 3 | Autogenerado desde `ManoObra.id` |
| OTRO | `OT-` + id zero-padded a 3 | Autogenerado desde `Otros.id` |
| Ítem sin BD (descripción manual, `recurso_id` vacío) | `—` | n/a |

**Por qué híbrido:** los materiales reales del cliente (Inelco, DICOL) ya tienen códigos de proveedor que él reconoce. Reemplazarlos por `MAT-001` perdería trazabilidad. Equipos/MO/Otros no tienen códigos del mundo real, así que autogenerar es seguro.

**Dónde se muestra:** la columna **ÍTEM** del APU impreso (hoy muestra el índice 1, 2, 3…) pasa a mostrar el código resuelto.

### 2. Factor de Desperdicio de Materiales

**Campo nuevo en la hoja `APU`:** `desperdicio_pct` (numérico, editable, default vacío/0).

**Aplicación:** afecta el `subtotal_materiales`, NO la cantidad individual de cada renglón. Esto preserva la trazabilidad — el ingeniero ve "necesito 12.5 m de tubería" y el sistema añade el % al subtotal final.

```
subtotal_materiales_base = Σ valor_parcial(materiales)
subtotal_materiales      = subtotal_materiales_base × (1 + desperdicio_pct/100)
```

**Render:** si `desperdicio_pct > 0`, el encabezado de la sección B muestra:
```
B — MATERIALES (Incluye factor de desperdicio del X%)
```
Si es 0 o vacío: encabezado normal sin el sufijo.

**Editable por APU.** No hay constante hardcodeada. El ingeniero llena el % según la complejidad de la actividad.

### 3. Herramienta Menor (HM)

**Campo nuevo en la hoja `APU`:** `herramienta_menor_pct` (numérico, editable, default vacío/0).

**Cálculo:**
```
HM_valor = subtotal_mano_obra × herramienta_menor_pct / 100
subtotal_equipos = Σ valor_parcial(equipos) + HM_valor
```

**Render:** si `herramienta_menor_pct > 0`, se inserta una línea automática al final de la sección A (Equipos), antes del subtotal:

```
A — EQUIPOS
ÍTEM     DESCRIPCIÓN                CANT.  REND.  P. UNIT.    VALOR PARCIAL
EQ-004   Dobladora Hidráulica       1.0    8.0    $45,000     $5,625
HM       Herramienta Menor (% MO)   —      —      subtotal_MO HM_valor   ← automática
                                                   SUBTOTAL   $6,780
```

Si es 0 o vacío: la línea no aparece y el cálculo no la incluye.

### 4. Encabezado informativo de Mano de Obra

La sección C siempre muestra: `C — MANO DE OBRA (Prestaciones Sociales del 54%)`.

Esto es texto fijo informativo. No hay campo nuevo. (Si en el futuro el cliente necesita distinguir 54% vs. 20% — algunos roles ya tienen 20% — se reabrirá el diseño.)

## Cambios técnicos

### Esquema (`setup.gs`)

**Hoja `APU`** — añadir 2 columnas al final del header existente:

```js
createSheet(ss, "APU", [
  "id", "codigo_item", "descripcion", "unidad", "cliente", "direccion",
  "actividad", "subtotal_equipos", "subtotal_materiales",
  "subtotal_mano_obra", "subtotal_otros", "costo_neto",
  "administracion_pct", "imprevistos_pct", "utilidad_pct", "iva_pct", "valor_total",
  "fecha",
  "desperdicio_pct", "herramienta_menor_pct"   // ← NUEVAS
]);
```

**Otras hojas:** sin cambios.

### Backend (`apu.gs`)

1. **`recalcularSubtotales(ss, apuId)`** — modificar el cálculo de `subtotal_materiales` y `subtotal_equipos` aplicando las nuevas fórmulas. Leer los `_pct` desde la fila del APU.

2. **`renderAPU(...)` / función de exportación** — en el bloque que pinta cada ítem (apu.gs:311 aprox), reemplazar `idx + 1` en columna 1 por la función `resolverCodigo(item)` que retorna el código según el tipo. En el encabezado de sección, añadir el sufijo de desperdicio para Materiales y el de prestaciones para MO. En la sección Equipos, insertar la fila HM antes del subtotal si aplica.

3. **`getAPU(apuId)` / `crearAPU(datos)` / `actualizarAPU(apuId, cambios)`** — incluir los nuevos campos en lectura/escritura.

4. **Función nueva `resolverCodigo(item)`** — devuelve el string según `item.tipo` y `item.recurso_id`. Para MATERIAL hace lookup a `Materiales.codigo`; para los otros aplica zero-pad.

### Frontend (`index.html`)

1. **Formulario APU** — añadir dos inputs numéricos en la sección de configuración del APU (junto a los campos AIU existentes):
   - `desperdicio_pct` con label "Desperdicio Materiales (%)"
   - `herramienta_menor_pct` con label "Herramienta Menor (% MO)"
2. **Eventos `onchange`** — disparar el recálculo del APU (igual que ya hace al cambiar admin/imprevistos/utilidad).
3. **Vista previa de items** — la tabla del APU en pantalla ya muestra una columna; añadir el código si está disponible para que el ingeniero vea lo mismo que verá en la impresión.

## Migración

**Cero migración de datos.** Las nuevas columnas quedan vacías en los APUs existentes → equivalente a `0%` → los cálculos no cambian para APUs previos. El primer `setupDatabase()` después del cambio:

- Si la hoja `APU` ya existe con el esquema viejo: agregar las 2 columnas nuevas al final con valores vacíos. Necesita lógica de migración en `setup.gs` (alternativamente, `createSheet` debe detectar columnas faltantes y añadirlas).

## Casos borde

- **Ítem con `recurso_id` vacío** (descripción manual escrita a mano): la columna ÍTEM muestra `—`. No bloquea nada.
- **`Materiales.codigo` vacío para un material** (BD mal poblada): la columna ÍTEM muestra `MAT-` + id como fallback.
- **`desperdicio_pct` o `herramienta_menor_pct` con valor negativo**: tratar como 0. Validar en el form (`min="0"`).
- **APU sin Mano de Obra pero con HM activa**: `HM_valor = 0`, la línea HM se muestra con valor 0 (o se omite si `HM_valor == 0`).

## Fuera de alcance

- Override de desperdicio por ítem (decisión: no, mantener simple).
- Cambiar el prefijo de los códigos APU (sigue siendo `codigo_item` libre).
- Soporte multi-prestaciones (algunos roles tienen 20% y otros 54% — el texto del encabezado dice "54%" estático; revisar después si el cliente lo pide).
- Cambios en el flujo de Cotizaciones (este spec es solo APU).

## Criterios de éxito

1. Un APU con `desperdicio_pct = 5` muestra "(Incluye factor de desperdicio del 5%)" en el encabezado de Materiales, y su subtotal_materiales = base × 1.05.
2. Un APU con `herramienta_menor_pct = 5` y subtotal_MO = $100,000 incluye una línea `HM` con valor $5,000 al final de Equipos.
3. Cada ítem de la sección Materiales muestra su código real de proveedor (AA1, TC15…) en la columna ÍTEM.
4. Cada ítem de Equipos/MO/Otros muestra `EQ-001`, `MO-001`, `OT-001` (autogenerados).
5. Los APUs creados antes del cambio siguen calculando y renderizando idénticamente (cero regresión).
