function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // === BASE DE DATOS DE PRECIOS ===
  createSheet(ss, "Equipos", [
    "id", "nombre", "tarifa_dia"
  ]);

  createSheet(ss, "Materiales", [
    "id", "codigo", "categoria", "nombre", "unidad",
    "precio_sin_iva", "precio_con_iva", "precio_2026",
    "proveedor", "fecha_actualizacion"
  ]);

  createSheet(ss, "ManoObra", [
    "id", "descripcion", "salario_mensual",
    "prestaciones_pct", "costo_dia"
  ]);

  createSheet(ss, "Otros", [
    "id", "descripcion", "unidad", "costo_unitario"
  ]);

  // === APU ===
  createSheet(ss, "APU", [
    "id", "codigo_item", "descripcion", "unidad", "cliente",
    "actividad", "subtotal_equipos", "subtotal_materiales",
    "subtotal_mano_obra", "subtotal_otros", "costo_neto", "fecha"
  ]);

  createSheet(ss, "APU_Items", [
    "id", "apu_id", "tipo", "recurso_id", "descripcion_manual",
    "cantidad", "rendimiento", "precio_unitario", "valor_parcial"
  ]);

  // === COTIZACIONES ===
  createSheet(ss, "Cotizaciones", [
    "id", "numero_oferta", "cliente", "direccion", "fecha",
    "valor_neto", "administracion_pct", "imprevistos_pct",
    "utilidad_pct", "valor_total", "aprobada", "notas"
  ]);

  createSheet(ss, "Cotizacion_Items", [
    "id", "cotizacion_id", "apu_id", "item_num", "descripcion",
    "unidad", "cantidad", "precio_apu", "valor_total"
  ]);

  seedMateriales(ss);
  SpreadsheetApp.flush();
  seedManoObra(ss);
  seedEquipos(ss);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    "✅ Base de datos creada y datos cargados.\n\n" +
    "Ahora corre la función «formatearHojas» para aplicar el formato visual."
  );
}
