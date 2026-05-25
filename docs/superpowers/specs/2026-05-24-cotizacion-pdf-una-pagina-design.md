# Spec: Cotización PDF — Una Sola Página

**Fecha:** 2026-05-24
**Estado:** Aprobado

## Problema

Al exportar cotizaciones largas (6-10 APUs con sub-ítems) el PDF resultante se divide en múltiples hojas. El cliente recibe 2-3 páginas cuando debería recibir siempre 1 página limpia.

## Solución

Dos cambios combinados en `src/cotizacion.gs`:

### 1. Layout compacto en `llenarHojaCotizacion`

Reducir alturas de fila y fuentes en toda la función para que el contenido ocupe menos espacio vertical de forma natural, minimizando cuánto tiene que escalar Google Sheets.

| Elemento | Altura actual | Altura nueva |
|---|---|---|
| Logo (fila 1) | 85px | 55px |
| Imagen del logo | 215×78 | 160×48 |
| Fila cliente/dirección | 34px | 22px |
| Fila número oferta | 20px | 14px |
| Encabezado tabla | 26px | 20px |
| Fila principal APU | 24px | 18px |
| Encabezado de sección (EQUIPOS, etc.) | 17px | 12px |
| Sub-ítems | variable | 13px |
| Subtotales de sección | 20px | 15px |
| Subtotal de APU | 20px | 15px |
| TOTAL COSTOS DIRECTOS | 24px | 18px |
| Filas AIU | 20px | 15px |
| VALOR TOTAL OFERTA | 28px | 20px |
| Espaciadores `r += 2` (antes de condiciones y firma) | 2 filas vacías | 1 fila vacía |

Las fuentes se reducen 1pt donde estaban en 10pt o más (encabezado tabla: 9→8, TOTAL COSTOS DIRECTOS: 10→9, VALOR TOTAL: 12→10).

### 2. Parámetro `scale=4` en URLs de exportación PDF

Reemplazar `fitw=true` por `scale=4` en ambas funciones que generan PDF:
- `enviarCotizacionEmail`
- `enviarYGuardarPDF`

`scale=4` = "Fit to page" de Google Sheets — escala el contenido proporcionalmente hasta que todo entre en exactamente 1 página. A diferencia de `fitw=true` (solo ajusta ancho), `scale=4` ajusta ambas dimensiones.

Reducir también los márgenes de `0.75` a `0.50` pulgadas en ambas funciones.

```
&scale=4&size=letter&portrait=true
&gridlines=false&printtitle=false&sheetnames=false
&top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50
```

## Archivos afectados

- `src/cotizacion.gs` — función `llenarHojaCotizacion` (layout) + `enviarCotizacionEmail` + `enviarYGuardarPDF` (URL)

## Lo que NO cambia

- La función `descargarPDFBase64` exporta `.xlsx`, no PDF — no se toca.
- La estructura de datos, cálculos, y lógica de negocio no cambian.
- El formato visual del documento se mantiene igual, solo más compacto.

## Resultado esperado

Con 6-10 APUs típicos (3-4 sub-ítems por sección), el PDF resultante cabe en 1 página carta. Para cotizaciones extremadamente largas (>10 APUs muy densos), el texto quedará pequeño pero siempre será 1 página.
