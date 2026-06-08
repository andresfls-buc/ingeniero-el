# Rediseño: Generación de Cotizaciones para Cliente

**Fecha:** 2026-06-08
**Estado:** Aprobado el diseño, pendiente plan de implementación

## Objetivo

Rediseñar el sitio de "clientes/cotizaciones" donde el ingeniero **arma una cotización**
escogiendo APUs que ya tiene creados. El módulo de APU ya está listo y **no se toca**.

La cotización es **una sola** (se elimina la idea de "interna vs cliente" como dos hojas
paralelas). El enredo y el bug actual (`doc=null` al "Pasar a Cotización Cliente") nacen de
mantener dos hojas sincronizadas: `Cotizacion_Items` y `Cotizacion_Cliente_Items`. El rediseño
usa **una sola fuente de verdad**.

## Principio clave

De cada APU, la cotización **solo trae 2 datos**: el **nombre** (descripción) y el **valor
unitario**. **Nunca** se muestra el interior del APU (ni mano de obra, ni equipos, ni
materiales, ni otros). Esos quedan ocultos dentro del APU.

El **valor unitario** de cada ítem = el `costo_neto` del APU, traído automáticamente al
agregarlo, **pero editable** en la cotización (para redondear o dar descuento) sin alterar el APU.

## Cómo funciona (flujo del usuario)

1. **Crear/abrir una cotización** y llenar datos del cliente (cliente, dirección/obra, objeto).
2. **Agregar APUs**: escoge de los APUs disponibles → entran trayendo nombre + valor unitario.
3. **Poner cantidad** a cada ítem → el sistema calcula `valor_total = cantidad × valor_unitario`.
4. **(Opcional) Agrupar en capítulos**: el usuario pone el número del capítulo (3, 4, …) y el
   sistema numera los ítems solo (3.1, 3.2, …). Sin capítulos → lista corrida 1, 2, 3.
5. **Cierre de costos** (AIU visible/desglosado):
   ```
   COSTO DIRECTO (suma de valores totales)
   + Administración (A%)
   + Imprevistos   (I%)
   + Utilidad      (U%)
   = SUBTOTAL
   + IVA (19%)
   = VALOR TOTAL DE LA OFERTA
   ```
6. **Entregar**: botón **descargar `.xlsx`** y botón **enviar PDF por correo** al cliente.

## Numeración (resuelve "no sé qué número vendrá")

- El **número del capítulo** lo decide el usuario (editable; por defecto sugiere el siguiente).
- El **sufijo del ítem** (`.1`, `.2`, …) lo asigna el sistema según el orden dentro del capítulo.
- Al reordenar o borrar un ítem, se renumera automáticamente.
- Sin capítulos: numeración corrida 1, 2, 3.

## Documento final (layout)

```
[EMPRESA / LOGO]                 OFERTA N° 2026-001
NIT: ___                         Fecha: 08/06/2026
CLIENTE: ___          OBRA/DIRECCIÓN: ___
OBJETO: Suministro e instalación de ___

ÍTEM  DESCRIPCIÓN          UND  CANT  V.UNIT   V.TOTAL
1.1   Tubería cobre 1"     ml   200  85.000   17.000.000
1.2   Codo 90° cobre 1"    un    40  12.000      480.000
                                  COSTO DIRECTO  17.480.000
                                  Administración (A%)  ___
                                  Imprevistos    (I%)  ___
                                  Utilidad       (U%)  ___
                                  SUBTOTAL             ___
                                  IVA (19%)            ___
                                  VALOR TOTAL          ___

Validez: 30 días · Forma de pago: ___ · Plazo: ___ · No incluye: ___
```

## Modelo de datos

Reutiliza las hojas existentes. Cambios mínimos:

**`Cotizaciones`** (ya existe — agregar 1 columna):
`id, numero_oferta, cliente, direccion, fecha, valor_neto, administracion_pct, imprevistos_pct,
utilidad_pct, iva_pct, valor_total, aprobada, notas, forma_pago, plazo_entrega, validez_oferta,
no_incluye` **+ `objeto`** (descripción de la obra).

**`Cotizacion_Items`** (ya existe — agregar 2 columnas):
`id, cotizacion_id, apu_id, item_num, descripcion, unidad, cantidad, precio_apu, valor_total`
**+ `capitulo_num`, `capitulo_nombre`**.

- `precio_apu` = valor unitario (traído del `costo_neto` del APU, editable).
- `item_num` = número compuesto calculado (`capitulo_num.sufijo`, ej. `3.1`).

**Datos de empresa** (oferente: nombre, NIT, logo, contacto) → en hoja `Configuracion`
(clave/valor), ya que son fijos del negocio.

**A deprecar:** la hoja `Cotizacion_Cliente_Items` y las funciones de `cotizacionCliente.gs`
que dependían de ella (origen del bug). Se reemplazan por operaciones sobre `Cotizacion_Items`.

## Componentes (backend `.gs`)

- **CRUD de cotización**: crear, abrir, actualizar datos de cabecera (cliente, objeto, A/I/U/IVA,
  condiciones), eliminar. (Reusar/limpiar lo de `cotizacion.gs`.)
- **CRUD de ítems**: agregar APU (trae nombre + valor unitario), editar cantidad/valor unitario,
  asignar capítulo, eliminar, reordenar. Recalcula `item_num` y totales.
- **Cálculo de totales**: `valor_total` por ítem → costo directo → AIU → IVA → valor total.
- **Render del documento**: una sola vista imprimible del documento del cliente (sin internos).
- **Export `.xlsx`**: genera archivo formateado y lo guarda en Drive / descarga.
- **Enviar PDF por correo**: genera PDF y lo envía al email del cliente desde la app.

## Frontend (`index.html`)

- Vista de **lista de cotizaciones** (crear nueva, abrir, eliminar).
- Vista de **edición de cotización**: datos de cabecera + tabla de ítems editable (agregar APU
  desde selector, cantidad, valor unitario, capítulo) + panel de cierre (A/I/U/IVA) con totales
  en vivo.
- Vista de **documento del cliente** (imprimible) + botones **Descargar .xlsx** y **Enviar PDF**.

## Errores y validaciones

- APU inexistente al agregar → mensaje claro, no romper.
- Cantidades/valores no numéricos → validar y avisar.
- Cotización sin ítems → no permitir exportar/enviar.
- Email inválido al enviar → validar antes de mandar.
- Retornos de `google.script.run` siempre objetos serializables simples (evitar el bug de
  serialización: nada de objetos complejos/circulares; si hay duda, validar en cliente).

## Fuera de alcance (YAGNI por ahora)

- Edición visual del logo/formato más allá de lo necesario para el documento estándar.
- Importar/exportar a formatos distintos de `.xlsx`/PDF.
- Flujo de aprobación del cliente dentro de la app.

## Datos de negocio que el usuario debe proveer (no técnicos)

- Nombre / NIT / logo / contacto de la empresa.
- Porcentajes típicos de Administración, Imprevistos, Utilidad.
- Condiciones comerciales por defecto (validez, forma de pago, plazo).
