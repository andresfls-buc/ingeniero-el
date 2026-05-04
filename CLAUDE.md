# Sistema APU - Ingeniero Mecánico

## Qué es este proyecto

Automatización del flujo de cotización de un ingeniero mecánico colombiano que trabaja con instalaciones de tubería (cobre, acero, gas). Actualmente hace todo a mano en Excel; este sistema lo reemplaza con Google Apps Script + Google Sheets.

## Stack

- **Google Apps Script** (`.gs`) — backend, lógica de negocio
- **Google Sheets** — base de datos y visualización
- **HTML/CSS/JS** (`index.html`) — frontend como sidebar/webapp de Google Sheets

## Flujo de trabajo del ingeniero

```
BD de Precios → APU (por actividad) → Cuadro General de Cotización → Registro de Ofertas
```

1. **BD de precios**: tablas dinámicas de materiales, equipos, mano de obra (actualizables)
2. **APU** (Análisis de Precios Unitarios): hoja por actividad con 4 secciones
3. **Cotización**: agrupa varios APUs con AIU (administración + imprevistos + utilidad)
4. **Registro de Ofertas**: historial con número, cliente, fecha, valor, aprobación (sí/no)

## Estructura del APU

Cada APU tiene datos generales + 4 secciones:

**Datos generales**: `codigo_item`, `unidad` (m², unidad, ml), `cliente/obra`, `actividad`

| Sección | Campos clave |
|---------|-------------|
| Equipos | nombre, tarifa_día (de BD), rendimiento (manual), valor_parcial |
| Materiales | descripción (de BD), unidad, precio_sin_iva (de BD), cantidad, valor_parcial |
| Mano de Obra | descripción (de BD), cantidad personas, costo_día (calculado), rendimiento/día, valor_parcial |
| Otros | transporte, izaje, etc. — campo abierto |

`valor_neto = subtotal_equipos + subtotal_materiales + subtotal_mano_obra + subtotal_otros`

El AIU (admin + imprevistos + utilidad) normalmente se aplica en el cuadro general, NO en el APU.

## Cálculo de Mano de Obra (Ley Colombiana)

```
salario_integral = salario_mensual * 1.54   (54% prestaciones: SGSSS, ARL, primas, dotación)
costo_dia = salario_integral / 30
```

- Salario mínimo 2025: $1,750,905 + aux. transporte $249,095 = **$2,000,000/mes**
- Ayudante: $2,000,000 × 1.54 / 30 = **$102,667/día**
- Ingeniero: $4,000,000 × 1.54 / 30 = **$205,333/día**

## Hojas de Google Sheets

### BD de Precios
| Hoja | Columnas clave |
|------|---------------|
| `Equipos` | id, nombre, tarifa_dia |
| `Materiales` | id, codigo, categoria, nombre, unidad, precio_sin_iva, precio_con_iva, precio_2026, proveedor, fecha_actualizacion |
| `ManoObra` | id, descripcion, salario_mensual, prestaciones_pct, costo_dia |
| `Otros` | id, descripcion, unidad, costo_unitario |

### APU
| Hoja | Columnas clave |
|------|---------------|
| `APU` | id, codigo_item, descripcion, unidad, cliente, actividad, subtotal_equipos, subtotal_materiales, subtotal_mano_obra, subtotal_otros, costo_neto, fecha |
| `APU_Items` | id, apu_id, tipo (EQUIPO/MATERIAL/MANO_OBRA/OTRO), recurso_id, descripcion_manual, cantidad, rendimiento, precio_unitario, valor_parcial |

### Cotizaciones
| Hoja | Columnas clave |
|------|---------------|
| `Cotizaciones` | id, numero_oferta, cliente, direccion, fecha, valor_neto, administracion_pct, imprevistos_pct, utilidad_pct, valor_total, aprobada, notas |
| `Cotizacion_Items` | id, cotizacion_id, apu_id, item_num, descripcion, unidad, cantidad, precio_apu, valor_total |

## Materiales precargados (seedData.gs)

138 materiales en 7 categorías del listado real del cliente:

| Categoría | Ejemplos |
|-----------|---------|
| Accesorio CU | ADAP H/M, CODO 45°/90°, TEE, UNION C/T, TAPON, RED BUS/COPA, UNIVERSAL |
| Tub CU Flex ACR | TC1/4 a TC10 (1/8" a 1 1/8") — precio por metro |
| Tub CU Rig K | TC11 a TC21 (1/4" a 3") — precio por metro, barra 6mt |
| Tub CU Rig L | TC22 a TC34 (1/4" a 6") — precio por metro, barra 6mt, tiene precio_2026 |
| Tub Acero | CED 40 y CED 10, 1" a 8", precio por metro (barra 5.80mt) |
| Aislamiento | 1/4" a 2 5/8" pared 1/2", por sección 1.82mt, tiene precio_2026 |
| Tub Galvanizada Gas | 1/2" a 2" CED 40 |
| Tub PE-AL-PE Gas | 1216, 1418, 1620, 2025 |

**Precios**: contado antes de IVA. IVA Colombia = 19%. `precio_2026` incluye alza del 8%.

## Archivos del proyecto

```
src/
├── Code.gs       — doGet(), entry point de la webapp
├── setup.gs      — setupDatabase(), crea/resetea todas las hojas
├── seedData.gs   — seedMateriales(), seedManoObra() con datos reales
├── helpers.gs    — createSheet() y utilidades
└── index.html    — UI frontend (en construcción)
```

## Lo que falta construir

- [ ] `apu.gs` — CRUD de APUs, cálculo automático de subtotales y costo neto
- [ ] `cotizacion.gs` — Ensamblar APUs en cotización, aplicar AIU, generar número de oferta
- [ ] `manoObra.gs` — Recalcular costo_dia cuando cambia salario mínimo
- [ ] `index.html` — UI completa: formulario APU (dropdown de BD), formulario cotización, historial de ofertas
- [ ] Exportar cotización terminada a formato Excel/PDF

## Convenciones

- Precios siempre en pesos colombianos (COP), sin decimales
- El APU usa `precio_sin_iva` de la BD de materiales
- El rendimiento lo ingresa el ingeniero manualmente por experiencia
- `tipo` en APU_Items: `"EQUIPO"`, `"MATERIAL"`, `"MANO_OBRA"`, `"OTRO"`
