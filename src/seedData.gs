// Carga todos los materiales del listado de tubería de cobre y accesorios.
// Se llama desde setupDatabase() en setup.gs
function seedMateriales(ss) {
  const sheet = ss.getSheetByName("Materiales");
  const fecha = new Date().toLocaleDateString("es-CO");
  const proveedor = "";

  // Formato: [codigo, categoria, nombre, unidad, precio_sin_iva, precio_con_iva, precio_2026]
  // precio_con_iva = null → se calcula automáticamente como precio_sin_iva * 1.19
  // precio_2026    = ""   → no disponible para ese ítem
  const data = [

    // ── ADAPTADORES HEMBRA ────────────────────────────────────────────────
    ["", "Accesorio CU", 'ADAP H CU 1/2" (5/8 R)',      "UND", 3621,  null, ""],
    ["", "Accesorio CU", 'ADAP H CU 3/4" (7/8 R)',      "UND", 6356,  null, ""],
    ["", "Accesorio CU", 'ADAP H CU 1" (1 1/8 R)',      "UND", 13054, null, ""],
    ["", "Accesorio CU", 'ADAP H CU 1 1/4" (1 3/8 R)',  "UND", 17208, null, ""],
    ["", "Accesorio CU", 'ADAP H CU 1 1/2" (1 5/8 R)',  "UND", 24579, null, ""],
    ["", "Accesorio CU", 'ADAP H CU 2" (2 1/8 R)',      "UND", 34753, null, ""],
    ["", "Accesorio CU", 'ADAP H CU 2 1/2" (2 5/8 R)',  "UND", 94121, null, ""],

    // ── ADAPTADORES MACHO ─────────────────────────────────────────────────
    ["", "Accesorio CU", 'ADAP M CU 1/4"',              "UND", 3666,  null, ""],
    ["", "Accesorio CU", 'ADAP M CU 3/8" (1/2 R)',      "UND", 5706,  null, ""],
    ["", "Accesorio CU", 'ADAP M CU 1/2" (5/8 R)',      "UND", 3578,  null, ""],
    ["", "Accesorio CU", 'ADAP M CU 3/4" (7/8 R)',      "UND", 6284,  null, ""],
    ["", "Accesorio CU", 'ADAP M CU 1" (1 1/8 R)',      "UND", 11276, null, ""],
    ["", "Accesorio CU", 'ADAP M CU 1 1/4" (1 3/8 R)',  "UND", 17009, null, ""],
    ["", "Accesorio CU", 'ADAP M CU 1 1/2" (1 5/8 R)',  "UND", 24229, null, ""],
    ["", "Accesorio CU", 'ADAP M CU 2" (2 1/8 R)',      "UND", 34481, null, ""],

    // ── CODOS 45° ─────────────────────────────────────────────────────────
    ["", "Accesorio CU", 'CODO 45° CU 1/4" (3/8 R)',    "UND", 1007, null, ""],
    ["", "Accesorio CU", 'CODO 45° CU 3/8" (1/2 R)',    "UND", 1558, null, ""],
    ["", "Accesorio CU", 'CODO 45° CU 1/2" (5/8 R)',    "UND", 1292, null, ""],
    ["", "Accesorio CU", 'CODO 45° CU 5/8" (3/4 R)',    "UND", 2756, null, ""],
    ["", "Accesorio CU", 'CODO 45° CU 3/4" (7/8 R)',    "UND", 3341, null, ""],
    ["", "Accesorio CU", 'CODO 45° CU 1" (1 1/8 R)',    "UND", 6300, null, ""],
    ["", "Accesorio CU", 'CODO 45° CU 1 1/4" (1 3/8 R)',"UND", 8660, null, ""],

    // ── CODOS 90° ─────────────────────────────────────────────────────────
    ["", "Accesorio CU", 'CODO 90° CU 1/8"',            "UND", 784,    null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 1/4" (3/8 R)',    "UND", 1055,   null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 3/8" (1/2 R)',    "UND", 1190,   null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 1/2" (5/8 R)',    "UND", 1504,   null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 5/8" (3/4 R)',    "UND", 2999,   null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 3/4" (7/8 R)',    "UND", 3358,   null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 1" (1 1/8 R)',    "UND", 6029,   null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 1 1/4" (1 3/8 R)',"UND", 9704,   null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 1 1/2" (1 5/8 R)',"UND", 17280,  null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 2" (2 1/8 R)',    "UND", 32891,  null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 2 1/2" (2 5/8 R)',"UND", 64211,  null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 4" (4 1/8 R)',    "UND", 230327, null, ""],
    ["", "Accesorio CU", 'CODO 90° CU 6"',              "UND", 839298, null, ""],

    // ── REDUCCIONES BUSHING ───────────────────────────────────────────────
    ["", "Accesorio CU", 'RED BUS CU 1" x 1/2" (1 1/8 x 5/8 R)',       "UND", 4091, null, ""],
    ["", "Accesorio CU", 'RED BUS CU 1" x 3/4" (1 1/8 x 7/8 R)',       "UND", 4458, null, ""],
    ["", "Accesorio CU", 'RED BUS CU 1 1/4" x 1" (1 3/8 x 1 1/8 R)',   "UND", 6715, null, ""],
    ["", "Accesorio CU", 'RED BUS CU 1 1/4" x 3/4" (1 3/8 x 7/8 R)',   "UND", 6101, null, ""],

    // ── REDUCCIONES COPA ──────────────────────────────────────────────────
    ["", "Accesorio CU", 'RED COPA CU 1/4" x 1/8" (3/8 x 2/8 R)',      "UND", 989,   null, ""],
    ["", "Accesorio CU", 'RED COPA CU 1/2" x 1/4" (5/8 x 3/8 R)',      "UND", 1388,  null, ""],
    ["", "Accesorio CU", 'RED COPA CU 1/2" x 3/8" (5/8 x 1/2 R)',      "UND", 1221,  null, ""],
    ["", "Accesorio CU", 'RED COPA CU 5/8" x 1/2" (3/4" x 5/8 R)',     "UND", 1899,  null, ""],
    ["", "Accesorio CU", 'RED COPA CU 5/8" x 3/8" (3/4" x 1/2 R)',     "UND", 1803,  null, ""],
    ["", "Accesorio CU", 'RED COPA CU 3/4" x 1/2" (7/8 x 5/8 R)',      "UND", 2083,  null, ""],
    ["", "Accesorio CU", 'RED COPA CU 1" x 1/2" (1 1/8 x 5/8 R)',      "UND", 3844,  null, ""],
    ["", "Accesorio CU", 'RED COPA CU 1" x 3/4" (1 1/8 x 7/8 R)',      "UND", 4386,  null, ""],
    ["", "Accesorio CU", 'RED COPA CU 1 1/4" x 1/2" (1 3/8x5/8 R)',    "UND", 6101,  null, ""],
    ["", "Accesorio CU", 'RED COPA CU 2" x 1 1/2" (2 1/8x1 5/8 R)',    "UND", 19413, null, ""],

    // ── TAPONES ───────────────────────────────────────────────────────────
    ["", "Accesorio CU", 'TAPON CU 1/4" (3/8 R)',  "UND", 833,   null, ""],
    ["", "Accesorio CU", 'TAPON CU 3/8" (1/2 R)',  "UND", 1161,  null, ""],
    ["", "Accesorio CU", 'TAPON CU 1/2" (5/8 R)',  "UND", 902,   null, ""],
    ["", "Accesorio CU", 'TAPON CU 3/4" (7/8 R)',  "UND", 1803,  null, ""],
    ["", "Accesorio CU", 'TAPON CU 1" (1 1/8 R)',  "UND", 4386,  null, ""],
    ["", "Accesorio CU", 'TAPON CU 1 1/2" (1 5/8 R)',"UND",11774, null, ""],
    ["", "Accesorio CU", 'TAPON CU 2" (2 1/8 R)',  "UND", 26231, null, ""],

    // ── TEES ──────────────────────────────────────────────────────────────
    ["", "Accesorio CU", 'TEE CU 1/4" (3/8 R)',        "UND", 2254,   null, ""],
    ["", "Accesorio CU", 'TEE CU 3/8" (1/2 R)',        "UND", 2234,   null, ""],
    ["", "Accesorio CU", 'TEE CU 1/2" (5/8 R)',        "UND", 2501,   null, ""],
    ["", "Accesorio CU", 'TEE CU 5/8" (3/4 R)',        "UND", 5784,   null, ""],
    ["", "Accesorio CU", 'TEE CU 3/4" (7/8 R)',        "UND", 6356,   null, ""],
    ["", "Accesorio CU", 'TEE CU 1" (1 1/8 R)',        "UND", 12104,  null, ""],
    ["", "Accesorio CU", 'TEE CU 1 1/4" (1 3/8 R)',    "UND", 17830,  null, ""],
    ["", "Accesorio CU", 'TEE CU 1 1/2" (1 5/8 R)',    "UND", 29334,  null, ""],
    ["", "Accesorio CU", 'TEE CU 2" (2 1/8 R)',        "UND", 60686,  null, ""],
    ["", "Accesorio CU", 'TEE CU 3" (3 1/8 R)',        "UND", 200484, null, ""],
    ["", "Accesorio CU", 'TEE CU 4" (4 1/8 R)',        "UND", 384401, null, ""],

    // ── TEES REDUCIDAS ────────────────────────────────────────────────────
    ["", "Accesorio CU", 'TEE RED CU 1/2" x 1/2" x 3/8"',              "UND", 2831, null, ""],
    ["", "Accesorio CU", 'TEE RED CU 3/4" x 3/4" x 1/2" (7/8x5/8 R)', "UND", 4370, null, ""],
    ["", "Accesorio CU", 'TEE RED CU 3/4" x 3/4" x 3/8" (7/8x1/2 R)', "UND", 6077, null, ""],

    // ── UNIONES ───────────────────────────────────────────────────────────
    ["", "Accesorio CU", 'UNION C/T CU 1/8"',          "UND", 263,   null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 1/4" (3/8 R)',  "UND", 572,   null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 3/8" (1/2 R)',  "UND", 813,   null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 1/2" (5/8 R)',  "UND", 947,   null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 5/8" (3/4 R)',  "UND", 1886,  null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 3/4" (7/8 R)',  "UND", 2345,  null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 1" (1 1/8 R)',  "UND", 4259,  null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 1 1/4" (1 3/8 R)',"UND",6244,  null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 1 1/2" (1 5/8 R)',"UND",10568, null, ""],
    ["", "Accesorio CU", 'UNION C/T CU 2" (2 1/8 R)',  "UND", 21677, null, ""],

    // ── UNIVERSALES ───────────────────────────────────────────────────────
    ["", "Accesorio CU", 'UNIVERSAL 1/2" CU (5/8 R)',  "UND", 8819,  null, ""],
    ["", "Accesorio CU", 'UNIVERSAL 3/4" CU (7/8 R)',  "UND", 22646, null, ""],
    ["", "Accesorio CU", 'UNIVERSAL 1" CU (1 1/8 R)',  "UND", 31976, null, ""],

    // ── TUBERÍA COBRE FLEXIBLE ACR (precio por metro) ─────────────────────
    ["TC1/4", "Tub CU Flex ACR", 'TUB CU FLEX ACR 1/8"',    "MT", 3801,  4523,  ""],
    ["TC2",   "Tub CU Flex ACR", 'TUB CU FLEX ACR 3/16"',   "MT", 4925,  5861,  ""],
    ["TC3",   "Tub CU Flex ACR", 'TUB CU FLEX ACR 1/4"',    "MT", 6643,  7906,  ""],
    ["TC4",   "Tub CU Flex ACR", 'TUB CU FLEX ACR 5/16"',   "MT", 9451,  11247, ""],
    ["TC5",   "Tub CU Flex ACR", 'TUB CU FLEX ACR 3/8"',    "MT", 11072, 13176, ""],
    ["TC6",   "Tub CU Flex ACR", 'TUB CU FLEX ACR 1/2"',    "MT", 15039, 17896, ""],
    ["TC7",   "Tub CU Flex ACR", 'TUB CU FLEX ACR 5/8"',    "MT", 21240, 25275, ""],
    ["TC8",   "Tub CU Flex ACR", 'TUB CU FLEX ACR 3/4"',    "MT", 25809, 30713, ""],
    ["TC9",   "Tub CU Flex ACR", 'TUB CU FLEX ACR 7/8"',    "MT", 39453, 46949, ""],
    ["TC10",  "Tub CU Flex ACR", 'TUB CU FLEX ACR 1 1/8"',  "MT", 57406, 68313, ""],

    // ── TUBERÍA COBRE RÍGIDA TIPO K (precio por metro, barra 6mt) ─────────
    ["TC11", "Tub CU Rig K", 'TUB CU RIG K 1/4" (3/8 R)',    "MT", 11925,  14190,  15326],
    ["TC12", "Tub CU Rig K", 'TUB CU RIG K 3/8" (1/2 R)',    "MT", 22122,  26326,  28432],
    ["TC13", "Tub CU Rig K", 'TUB CU RIG K 1/2" (5/8 R)',    "MT", 27592,  32834,  35461],
    ["TC14", "Tub CU Rig K", 'TUB CU RIG K 5/8" (3/4 R)',    "MT", 35269,  41970,  45328],
    ["TC15", "Tub CU Rig K", 'TUB CU RIG K 3/4" (7/8 R)',    "MT", 51414,  61182,  66077],
    ["TC16", "Tub CU Rig K", 'TUB CU RIG K 1" (1 1/8 R)',    "MT", 68999,  82108,  88677],
    ["TC17", "Tub CU Rig K", 'TUB CU RIG K 1 1/4" (1 3/8 R)',"MT", 87750,  104423, 112776],
    ["TC18", "Tub CU Rig K", 'TUB CU RIG K 1 1/2" (1 5/8 R)',"MT", 114750, 136552, 147477],
    ["TC19", "Tub CU Rig K", 'TUB CU RIG K 2" (2 1/8 R)',    "MT", 197950, 235560, 254405],
    ["TC20", "Tub CU Rig K", 'TUB CU RIG K 2 1/2" (2 5/8 R)',"MT", 293281, 349005, 376925],
    // TC21: precio barra $2,402,304 / 6mt = $400,384/mt (precio con IVA y 2026 no disponibles en listado)
    ["TC21", "Tub CU Rig K", 'TUB CU RIG K 3" (3 1/8 R)',    "MT", 400384, null,   ""],

    // ── TUBERÍA COBRE RÍGIDA TIPO L (precio por metro, barra 6mt) ─────────
    ["TC22", "Tub CU Rig L", 'TUB CU RIG L 1/4" (3/8 R)',    "MT", 9863,   11737,  12676],
    ["TC23", "Tub CU Rig L", 'TUB CU RIG L 3/8" (1/2 R)',    "MT", 15314,  18224,  19682],
    ["TC24", "Tub CU Rig L", 'TUB CU RIG L 1/2" (5/8 R)',    "MT", 21784,  25923,  27996],
    ["TC25", "Tub CU Rig L", 'TUB CU RIG L 5/8" (3/4 R)',    "MT", 29771,  35427,  38261],
    ["TC26", "Tub CU Rig L", 'TUB CU RIG L 3/4" (7/8 R)',    "MT", 35192,  41878,  45228],
    ["TC27", "Tub CU Rig L", 'TUB CU RIG L 1" (1 1/8 R)',    "MT", 51271,  61012,  65893],
    ["TC28", "Tub CU Rig L", 'TUB CU RIG L 1 1/4" (1 3/8 R)',"MT", 69417,  82606,  89214],
    ["TC29", "Tub CU Rig L", 'TUB CU RIG L 1 1/2" (1 5/8 R)',"MT", 89424,  106415, 114928],
    ["TC30", "Tub CU Rig L", 'TUB CU RIG L 2" (2 1/8 R)',    "MT", 156097, 185755, 200616],
    ["TC31", "Tub CU Rig L", 'TUB CU RIG L 2 1/2" (2 5/8 R)',"MT", 252331, 300274, 324296],
    ["TC32", "Tub CU Rig L", 'TUB CU RIG L 3" (3 1/8 R)',    "MT", 364482, 433733, 468432],
    ["TC33", "Tub CU Rig L", 'TUB CU RIG L 4" (4 1/8 R)',    "MT", 588000, 699720, 755698],
    ["TC34", "Tub CU Rig L", 'TUB CU RIG L 6" (6 1/8 R)',    "MT", 833000, 991271, 1070572],

    // ── TUBERÍA ACERO CARBONO (precio por metro, barra 5.80mt) ───────────
    ["", "Tub Acero", 'TUB ACER CARB PINTADA ROJA CED 40 1" x 5.80mt',     "MT", 15379,  18301,  ""],
    ["", "Tub Acero", 'TUB ACER CARB RANU ROJA CED 10 1 1/4" x 5.80mt',   "MT", 16595,  19748,  ""],
    ["", "Tub Acero", 'TUB ACER CARB RANU ROJA CED 10 1 1/2" x 5.80mt',   "MT", 19078,  22702,  ""],
    ["", "Tub Acero", 'TUB ACER CARB RANU ROJA CED 10 2" x 5.80mt',       "MT", 23868,  28403,  ""],
    ["", "Tub Acero", 'TUB ACER CARB RANU ROJA CED 10 2 1/2" x 5.80mt',   "MT", 31854,  37907,  ""],
    ["", "Tub Acero", 'TUB ACER CARB RANU ROJA CED 10 3" x 5.80mt',       "MT", 38840,  46220,  ""],
    ["", "Tub Acero", 'TUB ACER CARB RANU ROJA CED 10 4" x 5.80mt',       "MT", 52297,  62233,  ""],
    ["", "Tub Acero", 'TUB ACER CARB RANU ROJA CED 10 6" x 5.80mt',       "MT", 87283,  103867, ""],
    ["", "Tub Acero", 'TUB ACER CARB RANU ROJA CED 10 8" x 5.80mt',       "MT", 159419, 189709, ""],

    // ── AISLAMIENTO TÉRMICO (precio por sección de 1.82mt) ───────────────
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 1/4" PARED 1/2" x 1.82mt',   "UND", 2558,  3044,  3287],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 3/8" PARED 1/2" x 1.82mt',   "UND", 2805,  3338,  3605],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 1/2" PARED 1/2" x 1.82mt',   "UND", 3012,  3584,  3871],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 5/8" PARED 1/2" x 1.82mt',   "UND", 3342,  3977,  4295],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 3/4" PARED 1/2" x 1.82mt',   "UND", 3383,  4026,  4348],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 7/8" PARED 1/2" x 1.82mt',   "UND", 3672,  4370,  4719],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 1 1/8" PARED 1/2" x 1.82mt', "UND", 4084,  4860,  5249],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 1 3/8" PARED 1/2" x 1.82mt', "UND", 4538,  5400,  5832],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 1 5/8" PARED 1/2" x 1.82mt', "UND", 5280,  6284,  6786],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 2 1/8" PARED 1/2" x 1.82mt', "UND", 8265,  9836,  10622],
    ["", "Aislamiento", 'AISLAMIENTO TÉRMICO 2 5/8" PARED 1/2" x 1.82mt', "UND", 11435, 13608, 14696],
    ["", "Aislamiento", "CINTA FOAM",                                       "UND", 28025, null,  ""],

    // ── TUBERÍA GALVANIZADA GAS CED 40 ───────────────────────────────────
    ["", "Tub Galvanizada Gas", 'TUB GALVANIZADA GAS 1/2" CED 40',  "MT", 60774,  null, ""],
    ["", "Tub Galvanizada Gas", 'TUB GALVANIZADA GAS 3/4" CED 40',  "MT", 78278,  null, ""],
    ["", "Tub Galvanizada Gas", 'TUB GALVANIZADA GAS 1" CED 40',    "MT", 113170, null, ""],
    ["", "Tub Galvanizada Gas", 'TUB GALVANIZADA GAS 1 1/4" CED 40',"MT", 151610, null, ""],
    ["", "Tub Galvanizada Gas", 'TUB GALVANIZADA GAS 1 1/2" CED 40',"MT", 181496, null, ""],
    ["", "Tub Galvanizada Gas", 'TUB GALVANIZADA GAS 2" CED 40',    "MT", 239654, null, ""],

    // ── TUBERÍA PE-AL-PE GAS ──────────────────────────────────────────────
    ["", "Tub PE-AL-PE Gas", "TUB 1216 PE-AL-PE GAS", "MT", 1522, null, ""],
    ["", "Tub PE-AL-PE Gas", "TUB 1418 PE-AL-PE GAS", "MT", 1815, null, ""],
    ["", "Tub PE-AL-PE Gas", "TUB 1620 PE-AL-PE GAS", "MT", 2491, null, ""],
    ["", "Tub PE-AL-PE Gas", "TUB 2025 PE-AL-PE GAS", "MT", 4047, null, ""],
  ];

  const rows = data.map((m, i) => {
    const [codigo, categoria, nombre, unidad, precio_sin_iva, precio_con_iva, precio_2026] = m;
    const conIva = precio_con_iva !== null ? precio_con_iva : Math.round(precio_sin_iva * 1.19);
    return [i + 1, codigo, categoria, nombre, unidad, precio_sin_iva, conIva, precio_2026 !== "" ? precio_2026 : "", proveedor, fecha];
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// Carga los roles de mano de obra con cálculo automático de costo/día
// según ley colombiana: salario_integral = salario * (1 + prestaciones%) / 30
function seedEquipos(ss) {
  const sheet = ss.getSheetByName("Equipos");
  const equipos = [
    ["Antorcha Oxiacetilénica",               80000],
    ["Soldador de Arco Eléctrico",            60000],
    ["Dobladora de Tubería Manual",           15000],
    ["Dobladora de Tubería Hidráulica",       45000],
    ["Cortadora de Tubería",                   8000],
    ["Roscadora Eléctrica 1/2\" a 2\"",       80000],
    ["Taladro Percutor 1/2\"",                20000],
    ["Amoladora / Esmeriladora 4\"",          25000],
    ["Compresor de Aire 50 L",               100000],
    ["Bomba de Vacío 2 etapas",               60000],
    ["Manómetro Manifold (recarga gas)",      15000],
    ["Andamio Tubular (módulo 1.5 m)",        15000],
    ["Escalera Metálica 6 m",                 12000],
    ["Equipo de Soldadura TIG",              120000],
    ["Llave de Cadena 36\"",                   8000],
    ["Llave Stilson 18\"",                     5000],
    ["Detector de Fugas de Gas",              20000],
    ["Prensaflex / Prensa Hidráulica",       150000],
    ["Máquina de Soldar MIG/MAG",            100000],
    ["Rotomartillo SDS Plus",                 30000],
    ["Sierra Eléctrica para Metal",           40000],
    ["Nivel Láser",                           25000],
    ["Generador Eléctrico 5 kVA",             80000],
    ["Fluxómetro (medidor de caudal)",        35000],
    ["Cámara Termográfica",                  120000],
  ];
  const rows = equipos.map((e, i) => [i + 1, e[0], e[1]]);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedManoObra(ss) {
  const sheet = ss.getSheetByName("ManoObra");
  const PREST_PCT = 54; // 54% cubre seguridad social, ARL, primas, dotación

  // [descripcion, salario_mensual]
  const trabajadores = [
    ["Ayudante",              2000000], // Salario mínimo + aux. transporte
    ["Técnico",               2500000],
    ["Ductero",               3000000],
    ["Plomero",               3000000],
    ["Técnico Especializado", 3500000],
    ["Ingeniero",             4000000],
  ];

  const rows = trabajadores.map((t, i) => {
    const [descripcion, salario] = t;
    const costo_dia = Math.round(salario * (1 + PREST_PCT / 100) / 30);
    return [i + 1, descripcion, salario, PREST_PCT, costo_dia];
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
