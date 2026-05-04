function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheet(ss, "Equipos", ["id", "nombre", "tarifa_dia", "rendimiento"]);

  createSheet(ss, "Materiales", ["id", "nombre", "unidad", "precio_unitario"]);

  createSheet(ss, "ManoObra", [
    "id",
    "descripcion",
    "salario_mensual",
    "prestaciones",
    "costo_dia",
  ]);

  createSheet(ss, "Otros", ["id", "descripcion", "costo"]);

  createSheet(ss, "APU", [
    "id",
    "item",
    "unidad",
    "cliente",
    "actividad",
    "costo_total",
  ]);

  createSheet(ss, "APU_Items", [
    "id",
    "apu_id",
    "tipo",
    "recurso_id",
    "cantidad",
    "rendimiento",
  ]);

  createSheet(ss, "Ofertas", [
    "id",
    "apu_id",
    "cantidad",
    "costo_unitario",
    "total",
  ]);

  createSheet(ss, "Cotizaciones", [
    "id",
    "cliente",
    "fecha",
    "total",
    "estado",
  ]);
}
