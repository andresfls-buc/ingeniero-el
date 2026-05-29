# Diseño: Cotización para el cliente (documento limpio) + escalabilidad a 300+ ítems

**Fecha:** 2026-05-29
**Estado:** Aprobado para escribir plan de implementación

## Problema

El cliente del ingeniero revisó la app y señaló dos cosas:

1. **La cotización que se entrega al cliente final no debe mostrar el desglose** de equipos, materiales, mano de obra, desperdicio ni herramienta menor. En Colombia, la cotización/presupuesto de obra que ve el cliente final solo muestra, por cada ítem: número, descripción, unidad, cantidad, valor unitario y valor total. El desglose vive en el APU, que es un documento anexo y separado.

2. **La numeración de los ítems la impone quien pide la cotización** (cliente o licitación). Cada obra trae su propia numeración jerárquica por capítulos (una va de "1.2" a "4.3", otra de "4.1" a "8.0"). No se puede automatizar; el ingeniero solo copia el número que le dieron. La descripción también es texto libre dictado por la licitación (ej: "1.1 Localización y replanteo").

Aclaración posterior del cliente: **el documento detallado de hoy NO sobra**. Le sirve para auditoría interna — saber cuánto va a gastar en la obra. Por eso no se elimina: se separa en dos documentos.

Constraint adicional: una licitación puede tener **300+ ítems**. La app debe manejar ese volumen sin colgarse (Apps Script tiene límite de 6 minutos por ejecución).

## Investigación: formato colombiano

Confirmado con fuentes oficiales (ADR, IDU, guías de presupuesto):

- El documento que ve el cliente es un **cuadro de presupuesto** con columnas: `ÍTEM | DESCRIPCIÓN | UND | CANT. | VR. UNITARIO | VR. TOTAL`.
- Los ítems se numeran de forma **jerárquica por capítulos** (1.0 → 1.1, 1.2; 2.0 → 2.1...).
- El desglose (materiales/MO/equipos) NO se muestra; está en el APU anexo.
- Al final va el **AIU** (Administración, Imprevistos, Utilidad) y el IVA.

Fuentes:
- https://www.adr.gov.co/wp-content/uploads/2024/10/FORMATO-PRESUPUESTO-OBRA-CIVIL.pdf
- https://presucosto.com/guia-apu-colombia
- https://www.idu.gov.co/Archivos_Portal/Micrositios/Normograma/1/Dise%C3%B1o_de_Proyectos/04_Instructivos_Guias_cartillas/GUDP017_ELABORACION_PRESUPUESTO_CONTRATOS_OBRA_CONSULTORIA_INTERVENTORIA_V1.0.pdf

## Concepto central: una cotización, dos documentos

De los mismos datos (`Cotizaciones` + `Cotizacion_Items`) salen dos salidas distintas:

| Documento | Para quién | Contenido | Estado |
|-----------|-----------|-----------|--------|
| **Cotización interna** | Ingeniero / auditoría | Desglose completo: equipos, materiales, MO, desperdicio, HM, subtotales por sección | Ya existe (`llenarHojaCotizacion`). **Sin cambios.** |
| **Cotización cliente** | Cliente final / licitación | Tabla limpia de 6 columnas + AIU + IVA + condiciones + firma. Sin desglose. | **Nueva** (`llenarHojaCotizacionCliente`). |

## Decisiones de diseño (confirmadas con el usuario)

1. **Número de ítem:** campo de texto libre por ítem ("4.1", "8.0"). Sin lógica automática.
2. **Editable por ítem:** número de ítem + descripción. La cantidad ya se edita hoy. Valor unitario y unidad se toman del APU (no editables cuando el ítem viene de un APU).
3. **Orden en el documento del cliente:** ordenado por `item_num` con **orden numérico real** (4.1, 4.2, 5.0, 10.1), no alfabético. El documento interno mantiene su orden actual.
4. **Origen de líneas:** mixto. Cada línea puede venir de un APU (trae el valor unitario calculado) o ser una línea manual (descripción + unidad + cantidad + valor unitario escritos a mano, sin APU).
5. **Cantidad de ítems:** sin límite. UI con "+ Agregar ítem" y lista que crece sola.

## Cambios en los datos: `Cotizacion_Items`

La hoja ya tiene las columnas `item_num` y `descripcion`. Se **reutilizan** (no se rompe la estructura):

| Columna | Antes | Ahora |
|---------|-------|-------|
| `item_num` | entero automático (1, 2, 3...) | **texto libre editable** ("1.1", "4.2", "8.0") |
| `descripcion` | copia de la descripción del APU | **editable**; por defecto la del APU, el ingeniero la puede reescribir |
| `apu_id` | siempre apunta a un APU | **vacío permitido** para líneas manuales |
| `unidad` | de la BD/APU | escrita a mano cuando la línea es manual |
| `cantidad` | editable (ya) | igual |
| `precio_apu` | `costo_neto` del APU | valor unitario escrito a mano cuando la línea es manual |
| `valor_total` | `cantidad * precio_apu` | igual |

Migración: las cotizaciones existentes tienen `item_num` numérico. Eso sigue siendo texto válido; no requieren conversión. El campo simplemente pasa a ser editable.

## Componentes

### 1. Backend — `cotizacion.gs`

**`llenarHojaCotizacionCliente(sheet, cot)`** (nueva)
- Qué hace: renderiza el cuadro de presupuesto limpio para el cliente final.
- Entrada: hoja temporal + objeto cotización con sus ítems (sin sub-ítems de APU).
- Render en bloque: arma una matriz 2D completa (encabezado, filas de ítems, AIU, IVA, total) y la escribe con un solo `setValues`. Formatos (moneda, negritas, bordes, fondos) se aplican por rangos agrupados con `RangeList`, no celda por celda.
- Ítems ordenados por `item_num` usando una función de orden numérico jerárquico (ver helper abajo).
- Reusa: encabezado (logo/empresa/cliente), bloque AIU+IVA, condiciones comerciales y firma del `llenarHojaCotizacion` actual — extraer esas partes a helpers compartidos para no duplicar.

**`getCotizacionCliente(cotId)`** (nueva, liviana)
- Lee `Cotizaciones` (1 fila) + `Cotizacion_Items` (filtrado por cotId) en **una sola lectura cada una**.
- NO carga `APU_Items` ni la BD de materiales (no hay desglose). Esto es lo que la hace escalar a 300+ ítems.
- Normaliza los % AIU/IVA igual que `getCotizacionCompleta`.

**`compararItemNum(a, b)`** (helper de orden)
- Convierte "4.1", "10.2" en algo comparable numéricamente por segmentos: split por ".", comparar cada segmento como número. "4.2" < "10.1" < "10.2".
- Las líneas sin número van al final.

**`agregarLineaManual(cotId, datos)`** (nueva)
- Inserta una fila en `Cotizacion_Items` con `apu_id` vacío y `descripcion`/`unidad`/`cantidad`/`precio_apu` escritos a mano.
- Calcula `valor_total` y llama `recalcularCotizacion` una vez.

**`actualizarItemCotizacion(itemId, cambios)`** (extiende lo existente)
- Hoy solo se puede cambiar la cantidad (`actualizarCantidadItem`). Se generaliza para también editar `item_num` y `descripcion` (y, en líneas manuales, `unidad`/`precio_apu`).

**Operación por lote (escalabilidad):**
- `agregarVariasLineas(cotId, lineas[])`: inserta N filas con un solo `getRange().setValues()` y recalcula totales una sola vez. Evita 200 viajes a la hoja para una licitación grande.

**Exports:**
- Se agrega un parámetro `tipo` ("cliente" | "interna") a las funciones de export existentes (`descargarPDFBase64`, `exportarCotizacionPDF`, `enviarCotizacionEmail`, `enviarYGuardarPDF`). Según el `tipo`, cada una llama a `llenarHojaCotizacionCliente` o a `llenarHojaCotizacion`.
- Por defecto (`tipo` omitido) el email y la descarga usan el documento **cliente**.

### 2. Frontend — `index.html`

- En la pantalla de cotización, por cada ítem: campo de texto chico para `item_num`, campo para `descripcion` editable, y los ya existentes (cantidad).
- Botón "+ Agregar ítem" con dos modos: **"Desde APU"** (selector actual) y **"Manual"** (formulario: descripción, unidad, cantidad, valor unitario).
- Dos botones de export separados y rotulados: **"Cotización para el cliente"** y **"Cotización interna (auditoría)"**.
- Lista larga eficiente: no recargar los 300 ítems en cada edición; actualizar solo la fila tocada.

## Rendimiento (requisito de primer orden: 300+ ítems)

1. **Render en bloque.** El documento del cliente se arma en una matriz en memoria y se escribe con un único `setValues`; los formatos se agrupan por rangos (`RangeList`). Nada de `setValue`/`setBackground` por celda en un bucle de 300 filas.
2. **Lectura única.** `getCotizacionCliente` lee cada hoja una sola vez y no toca `APU_Items` ni la BD.
3. **Escrituras por lote.** Agregar/actualizar varios ítems en una sola operación de hoja; recalcular totales una vez al final.

## Lo que NO cambia

- APU completo (`apu.gs`), BD, cálculo de subtotales, desperdicio, HM, prestaciones.
- `llenarHojaCotizacion` (documento interno de auditoría) — intacto.
- Export de APU (`exportarAPUaDrive`, `descargarAPUBase64`).

## Criterios de aceptación

- [ ] El documento del cliente muestra solo: Nº ítem, descripción, unidad, cantidad, vr. unitario, vr. total + AIU + IVA. Sin desglose.
- [ ] El ingeniero puede escribir libremente el número de ítem y la descripción de cada línea.
- [ ] Se puede agregar una línea manual (sin APU) escribiendo descripción + unidad + cantidad + valor unitario.
- [ ] Los ítems del documento del cliente salen ordenados por número de ítem en orden numérico real.
- [ ] El documento interno de auditoría sigue funcionando igual que hoy.
- [ ] Una cotización de 300 ítems genera el documento del cliente sin superar el límite de tiempo de Apps Script.
