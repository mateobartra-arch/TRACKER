// ============================================================
// 🚀 MISAGI DRIVER TRACKING — Google Apps Script COMPLETO
// ============================================================
// INSTRUCCIONES:
// 1. Crea un Google Sheet NUEVO y VACÍO
// 2. Ve a Extensiones → Apps Script
// 3. Pega TODO este código
// 4. Guarda y ejecuta la función: crearHojaDesde0()
// 5. Google pedirá permisos → Acepta todo
// 6. ¡LISTO! La hoja se crea sola con todo el formato
// ============================================================

// =============================================
// 🎨 PALETA DE COLORES MISAGI (basada en el logo)
// =============================================
const COLORES = {
  // Marca MISAGI
  verde_principal: '#1B7340',  // Verde esmeralda del logo
  verde_claro: '#28A060',  // Verde más claro
  verde_suave: '#E8F5E9',  // Verde muy suave (fondos)
  azul_navy: '#1B3A6B',  // Azul oscuro del texto del logo
  azul_medio: '#1E4D8C',  // Azul medio
  azul_claro: '#E3F2FD',  // Azul suave (fondos)

  // UI Base
  blanco: '#FFFFFF',
  gris_fondo: '#F8FAFB',
  gris_claro: '#E8ECF0',
  gris_medio: '#94A3B8',
  gris_oscuro: '#475569',
  negro_texto: '#1E293B',

  // Operaciones (según codificación de imagen)
  hudbay: '#8B4513',  // Marrón tierra
  hudbay_claro: '#F5E6D3',
  toquepala: '#228B22',  // Verde bosque
  toquepala_claro: '#E8F5E9',
  cuajone: '#FFD700',  // Dorado
  cuajone_claro: '#FFFDE7',
  hierro: '#1E90FF',  // Azul
  hierro_claro: '#E3F2FD',

  // Estados
  descanso: '#CBD5E1',  // Gris suave
  descanso_texto: '#64748B',
  parado: '#FCD34D',  // Amarillo
  parado_texto: '#92400E',
  feriado: '#C0C0C0',  // Gris/Plata
  media: '#94A3B8',  // Gris
};

// =============================================
// ⚙️ CONFIGURACIÓN
// =============================================
const CONFIG = {
  NOMBRE_HOJA: 'ROSTER',
  FILA_TITULO: 1,
  FILA_MESES: 3,
  FILA_FECHAS: 4,
  FILA_DIAS_SEMANA: 5,
  FILA_INICIO_DATOS: 6,
  COL_NOMBRE: 1,
  COL_CARGO: 2,
  COL_INICIO_FECHAS: 3,

  // Meses a generar (Enero 2026 → Diciembre 2026)
  FECHA_INICIO: new Date(2026, 0, 1),   // 1 Enero 2026
  FECHA_FIN: new Date(2026, 11, 31),     // 31 Diciembre 2026

  // Conductores iniciales
  CONDUCTORES: [
    { nombre: 'MAMANI LOPEZ YURY', cargo: 'Conductor' },
    { nombre: 'MAMANI CARRIZALES NÉSTOR', cargo: 'Conductor' },
    { nombre: 'KANA HUALLPA RENE', cargo: 'Conductor' },
    { nombre: 'ARAPA CHOQUEHUANCA LUIS MIGUEL', cargo: 'Conductor' },
    { nombre: 'FARFÁN VICTOR', cargo: 'Conductor' },
    { nombre: 'PAREDES QUISPE JORGE NARCISO', cargo: 'Conductor' },
    { nombre: '', cargo: '' },
    { nombre: '', cargo: '' },
    { nombre: '', cargo: '' },
    { nombre: '', cargo: '' },
  ],

  // Codificación de operaciones
  CODIGOS: {
    // --- HUDBAY (marrones/tierra) ---
    'H1': { color: '#8B4513', texto: '#FFFFFF', nombre: 'Hudbay Día 1', grupo: 'HUDBAY' },
    'H2': { color: '#A0522D', texto: '#FFFFFF', nombre: 'Hudbay Día 2', grupo: 'HUDBAY' },
    'H3': { color: '#CD853F', texto: '#FFFFFF', nombre: 'Hudbay Día 3', grupo: 'HUDBAY' },
    // --- TOQUEPALA (verdes) ---
    'T1': { color: '#228B22', texto: '#FFFFFF', nombre: 'Toquepala Día 1', grupo: 'TOQUEPALA' },
    'T2': { color: '#32CD32', texto: '#FFFFFF', nombre: 'Toquepala Día 2', grupo: 'TOQUEPALA' },
    // --- CUAJONE (dorado/naranja) ---
    'C1': { color: '#FFD700', texto: '#000000', nombre: 'Cuajone Día 1', grupo: 'CUAJONE' },
    'C2': { color: '#FFA500', texto: '#000000', nombre: 'Cuajone Día 2', grupo: 'CUAJONE' },
    // --- HIERRO / ACERO (azules) ---
    'A1': { color: '#1E90FF', texto: '#FFFFFF', nombre: 'Hierro Día 1', grupo: 'HIERRO' },
    'A2': { color: '#00BFFF', texto: '#000000', nombre: 'Hierro Día 2', grupo: 'HIERRO' },
    'A3': { color: '#87CEEB', texto: '#000000', nombre: 'Hierro Día 3', grupo: 'HIERRO' },
    // --- ESTADOS ---
    'D': { color: '#E2E8F0', texto: '#64748B', nombre: 'Descanso', grupo: 'DESCANSO' },
    'd': { color: '#E2E8F0', texto: '#64748B', nombre: 'Descanso', grupo: 'DESCANSO' },
    'P': { color: '#FEF08A', texto: '#92400E', nombre: 'Parado', grupo: 'PARADO' },
    'p': { color: '#FEF08A', texto: '#92400E', nombre: 'Parado', grupo: 'PARADO' },
    'M/D': { color: '#CBD5E1', texto: '#475569', nombre: 'Media Jornada', grupo: 'MEDIA' },
    'D/M': { color: '#CBD5E1', texto: '#475569', nombre: 'Media Jornada', grupo: 'MEDIA' },
    'FERIADO': { color: '#C0C0C0', texto: '#333333', nombre: 'Feriado', grupo: 'FERIADO' },
    // ➕ AGREGA NUEVAS OPERACIONES AQUÍ
  }
};

// =============================================
// 🏗️ CREAR HOJA DESDE CERO
// =============================================
function crearHojaDesde0() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Eliminar hoja existente si existe
  let hoja = ss.getSheetByName(CONFIG.NOMBRE_HOJA);
  if (hoja) {
    ss.deleteSheet(hoja);
  }

  hoja = ss.insertSheet(CONFIG.NOMBRE_HOJA);

  // Mover al inicio
  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(1);

  // ---- GENERAR FECHAS ----
  const fechas = [];
  let current = new Date(CONFIG.FECHA_INICIO);
  while (current <= CONFIG.FECHA_FIN) {
    fechas.push(new Date(current));
    current.setDate(current.getDate() + 1);
  }

  const totalCols = CONFIG.COL_INICIO_FECHAS + fechas.length - 1;
  const totalFilas = CONFIG.FILA_INICIO_DATOS + CONFIG.CONDUCTORES.length + 10;

  // Expandir hoja
  if (hoja.getMaxColumns() < totalCols) {
    hoja.insertColumnsAfter(hoja.getMaxColumns(), totalCols - hoja.getMaxColumns());
  }
  if (hoja.getMaxRows() < totalFilas) {
    hoja.insertRowsAfter(hoja.getMaxRows(), totalFilas - hoja.getMaxRows());
  }

  // ================================================
  // FILA 1: TÍTULO PRINCIPAL
  // ================================================
  // Se divide la merge en 2 partes para no cruzar la columna congelada (col 2)
  hoja.setRowHeight(1, 50);

  // Parte izquierda (cols 1-2)
  const tituloIzq = hoja.getRange(1, 1, 1, 2);
  tituloIzq.merge();
  tituloIzq.setBackground(COLORES.azul_navy);
  tituloIzq.setFontColor(COLORES.blanco);
  tituloIzq.setFontSize(12);
  tituloIzq.setFontWeight('bold');
  tituloIzq.setFontFamily('Arial');
  tituloIzq.setVerticalAlignment('middle');
  tituloIzq.setValue('🚛 MISAGI');

  // Parte derecha (cols 3-totalCols)
  const tituloDer = hoja.getRange(1, CONFIG.COL_INICIO_FECHAS, 1, totalCols - CONFIG.COL_INICIO_FECHAS + 1);
  tituloDer.merge();
  tituloDer.setValue('ROSTER DE SEGUIMIENTO DE CONDUCTORES');
  tituloDer.setBackground(COLORES.azul_navy);
  tituloDer.setFontColor(COLORES.blanco);
  tituloDer.setFontSize(16);
  tituloDer.setFontWeight('bold');
  tituloDer.setFontFamily('Arial');
  tituloDer.setVerticalAlignment('middle');
  tituloDer.setHorizontalAlignment('center');

  // ================================================
  // FILA 2: SUBTÍTULO
  // ================================================
  hoja.setRowHeight(2, 28);

  // Parte izquierda (cols 1-2)
  const subIzq = hoja.getRange(2, 1, 1, 2);
  subIzq.merge();
  subIzq.setBackground(COLORES.verde_principal);
  subIzq.setFontColor(COLORES.blanco);
  subIzq.setFontSize(9);
  subIzq.setFontFamily('Arial');
  subIzq.setVerticalAlignment('middle');
  subIzq.setValue('Tracking');

  // Parte derecha (cols 3-totalCols)
  const subDer = hoja.getRange(2, CONFIG.COL_INICIO_FECHAS, 1, totalCols - CONFIG.COL_INICIO_FECHAS + 1);
  subDer.merge();
  subDer.setValue('Control de Asistencia y Operaciones  |  ' +
    CONFIG.FECHA_INICIO.toLocaleDateString('es-PE', { month: 'long', year: 'numeric' }) + ' → ' +
    CONFIG.FECHA_FIN.toLocaleDateString('es-PE', { month: 'long', year: 'numeric' }));
  subDer.setBackground(COLORES.verde_principal);
  subDer.setFontColor(COLORES.blanco);
  subDer.setFontSize(10);
  subDer.setFontFamily('Arial');
  subDer.setVerticalAlignment('middle');
  subDer.setHorizontalAlignment('center');

  // ================================================
  // FILA 3: NOMBRES DE LOS MESES
  // ================================================
  hoja.setRowHeight(3, 26);
  const mesesNombres = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
    'JULIO', 'AGOSTO', 'SETIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

  // Encabezados de columnas fijas
  hoja.getRange(3, 1).setValue('');
  hoja.getRange(3, 2).setValue('');

  let mesAnterior = -1;
  let mesInicioCol = CONFIG.COL_INICIO_FECHAS;

  for (let i = 0; i < fechas.length; i++) {
    const col = CONFIG.COL_INICIO_FECHAS + i;
    const mes = fechas[i].getMonth();

    if (mes !== mesAnterior && mesAnterior !== -1) {
      // Merge y escribir mes anterior
      if (col - mesInicioCol > 1) {
        const rangoMes = hoja.getRange(3, mesInicioCol, 1, col - mesInicioCol);
        rangoMes.merge();
      }
      hoja.getRange(3, mesInicioCol).setValue(mesesNombres[mesAnterior]);
      hoja.getRange(3, mesInicioCol).setHorizontalAlignment('center');

      // Alternar colores verde/azul por mes
      const colorMes = (mesAnterior % 2 === 0) ? COLORES.azul_navy : COLORES.verde_principal;
      hoja.getRange(3, mesInicioCol, 1, col - mesInicioCol).setBackground(colorMes);
      hoja.getRange(3, mesInicioCol, 1, col - mesInicioCol).setFontColor(COLORES.blanco);
      hoja.getRange(3, mesInicioCol, 1, col - mesInicioCol).setFontWeight('bold');
      hoja.getRange(3, mesInicioCol, 1, col - mesInicioCol).setFontSize(10);

      mesInicioCol = col;
    }
    mesAnterior = mes;
  }
  // Último mes
  const ultimaColFecha = CONFIG.COL_INICIO_FECHAS + fechas.length - 1;
  if (ultimaColFecha - mesInicioCol >= 0) {
    const rangoMes = hoja.getRange(3, mesInicioCol, 1, ultimaColFecha - mesInicioCol + 1);
    rangoMes.merge();
    hoja.getRange(3, mesInicioCol).setValue(mesesNombres[mesAnterior]);
    hoja.getRange(3, mesInicioCol).setHorizontalAlignment('center');
    const colorMes = (mesAnterior % 2 === 0) ? COLORES.azul_navy : COLORES.verde_principal;
    rangoMes.setBackground(colorMes);
    rangoMes.setFontColor(COLORES.blanco);
    rangoMes.setFontWeight('bold');
    rangoMes.setFontSize(10);
  }

  // ================================================
  // FILA 4: FECHAS (día/mes)
  // ================================================
  hoja.setRowHeight(4, 24);
  hoja.getRange(4, 1).setValue('CONDUCTOR');
  hoja.getRange(4, 1).setFontWeight('bold');
  hoja.getRange(4, 1).setBackground(COLORES.azul_navy);
  hoja.getRange(4, 1).setFontColor(COLORES.blanco);
  hoja.getRange(4, 1).setFontSize(10);

  hoja.getRange(4, 2).setValue('CARGO');
  hoja.getRange(4, 2).setFontWeight('bold');
  hoja.getRange(4, 2).setBackground(COLORES.azul_navy);
  hoja.getRange(4, 2).setFontColor(COLORES.blanco);
  hoja.getRange(4, 2).setFontSize(10);

  for (let i = 0; i < fechas.length; i++) {
    const col = CONFIG.COL_INICIO_FECHAS + i;
    const f = fechas[i];
    hoja.getRange(4, col).setValue(f.getDate());
    hoja.getRange(4, col).setHorizontalAlignment('center');
    hoja.getRange(4, col).setFontSize(9);
    hoja.getRange(4, col).setFontWeight('bold');

    // Fondo según día de semana
    const dow = f.getDay();
    if (dow === 0) { // Domingo
      hoja.getRange(4, col).setBackground('#FEE2E2');
      hoja.getRange(4, col).setFontColor('#DC2626');
    } else if (dow === 6) { // Sábado
      hoja.getRange(4, col).setBackground('#FEF3C7');
      hoja.getRange(4, col).setFontColor('#92400E');
    } else {
      hoja.getRange(4, col).setBackground(COLORES.azul_claro);
      hoja.getRange(4, col).setFontColor(COLORES.azul_navy);
    }
  }

  // ================================================
  // FILA 5: DÍAS DE LA SEMANA
  // ================================================
  hoja.setRowHeight(5, 20);
  const diasSemana = ['D', 'L', 'M', 'X', 'J', 'V', 'S'];

  hoja.getRange(5, 1).setValue('');
  hoja.getRange(5, 1).setBackground(COLORES.gris_claro);
  hoja.getRange(5, 2).setValue('');
  hoja.getRange(5, 2).setBackground(COLORES.gris_claro);

  for (let i = 0; i < fechas.length; i++) {
    const col = CONFIG.COL_INICIO_FECHAS + i;
    const dow = fechas[i].getDay();
    hoja.getRange(5, col).setValue(diasSemana[dow]);
    hoja.getRange(5, col).setHorizontalAlignment('center');
    hoja.getRange(5, col).setFontSize(8);
    hoja.getRange(5, col).setFontColor(COLORES.gris_medio);
    hoja.getRange(5, col).setBackground(COLORES.gris_fondo);

    if (dow === 0) {
      hoja.getRange(5, col).setFontColor('#DC2626');
      hoja.getRange(5, col).setBackground('#FFF5F5');
    }
  }

  // ================================================
  // CONDUCTORES
  // ================================================
  for (let i = 0; i < CONFIG.CONDUCTORES.length; i++) {
    const fila = CONFIG.FILA_INICIO_DATOS + i;
    hoja.setRowHeight(fila, 28);

    // Nombre
    const cNombre = hoja.getRange(fila, 1);
    cNombre.setValue(CONFIG.CONDUCTORES[i].nombre);
    cNombre.setFontWeight('bold');
    cNombre.setFontSize(9);
    cNombre.setFontColor(COLORES.negro_texto);
    cNombre.setVerticalAlignment('middle');

    // Cargo
    const cCargo = hoja.getRange(fila, 2);
    cCargo.setValue(CONFIG.CONDUCTORES[i].cargo);
    cCargo.setFontSize(8);
    cCargo.setFontColor(COLORES.gris_medio);
    cCargo.setVerticalAlignment('middle');

    // Fondo alternado para filas
    const bgFila = (i % 2 === 0) ? COLORES.blanco : COLORES.gris_fondo;
    hoja.getRange(fila, 1, 1, totalCols).setBackground(bgFila);

    // Borde izquierdo verde en la celda del nombre
    cNombre.setBorder(null, null, null, null, null, null, null, null);
  }

  // ================================================
  // FORMATO GENERAL
  // ================================================
  // Ancho de columnas
  hoja.setColumnWidth(1, 230);  // Nombre
  hoja.setColumnWidth(2, 80);   // Cargo

  for (let c = CONFIG.COL_INICIO_FECHAS; c <= totalCols; c++) {
    hoja.setColumnWidth(c, 32);
  }

  // Congelar filas de encabezado y columnas de nombre
  hoja.setFrozenRows(5);
  hoja.setFrozenColumns(2);

  // Bordes sutiles
  const rangoDatos = hoja.getRange(CONFIG.FILA_INICIO_DATOS, 1,
    CONFIG.CONDUCTORES.length, totalCols);
  rangoDatos.setBorder(true, true, true, true, true, true,
    COLORES.gris_claro, SpreadsheetApp.BorderStyle.SOLID);

  // ================================================
  // LEYENDA (debajo de los datos)
  // ================================================
  const filaLeyenda = CONFIG.FILA_INICIO_DATOS + CONFIG.CONDUCTORES.length + 2;

  hoja.getRange(filaLeyenda, 1, 1, 2).merge();
  hoja.getRange(filaLeyenda, 1).setValue('📋  LEYENDA');
  hoja.getRange(filaLeyenda, 1).setBackground(COLORES.azul_navy);
  hoja.getRange(filaLeyenda, 1).setFontColor(COLORES.blanco);
  hoja.getRange(filaLeyenda, 1).setFontWeight('bold');
  hoja.getRange(filaLeyenda, 1).setFontSize(10);
  hoja.getRange(filaLeyenda, CONFIG.COL_INICIO_FECHAS, 1, 6).merge();
  hoja.getRange(filaLeyenda, CONFIG.COL_INICIO_FECHAS).setValue('CÓDIGOS DE OPERACIÓN');
  hoja.getRange(filaLeyenda, CONFIG.COL_INICIO_FECHAS).setBackground(COLORES.azul_navy);
  hoja.getRange(filaLeyenda, CONFIG.COL_INICIO_FECHAS).setFontColor(COLORES.blanco);
  hoja.getRange(filaLeyenda, CONFIG.COL_INICIO_FECHAS).setFontWeight('bold');
  hoja.getRange(filaLeyenda, CONFIG.COL_INICIO_FECHAS).setFontSize(10);
  hoja.setRowHeight(filaLeyenda, 30);

  const leyendaData = [
    ['H1, H2, H3', 'HUDBAY', '#8B4513', '#FFFFFF'],
    ['T1, T2', 'TOQUEPALA', '#228B22', '#FFFFFF'],
    ['C1, C2', 'CUAJONE', '#FFD700', '#000000'],
    ['A1, A2, A3', 'HIERRO / ACERO', '#1E90FF', '#FFFFFF'],
    ['D', 'DESCANSO', '#E2E8F0', '#64748B'],
    ['P', 'PARADO / DISPONIBLE', '#FEF08A', '#92400E'],
    ['M/D', 'MEDIA JORNADA', '#CBD5E1', '#475569'],
    ['FERIADO', 'FERIADO', '#C0C0C0', '#333333'],
  ];

  for (let i = 0; i < leyendaData.length; i++) {
    const fl = filaLeyenda + 1 + i;
    hoja.getRange(fl, 1).setValue(leyendaData[i][0]);
    hoja.getRange(fl, 1).setBackground(leyendaData[i][2]);
    hoja.getRange(fl, 1).setFontColor(leyendaData[i][3]);
    hoja.getRange(fl, 1).setFontWeight('bold');
    hoja.getRange(fl, 1).setHorizontalAlignment('center');
    hoja.getRange(fl, CONFIG.COL_INICIO_FECHAS, 1, 3).merge();
    hoja.getRange(fl, CONFIG.COL_INICIO_FECHAS).setValue(leyendaData[i][1]);
    hoja.getRange(fl, CONFIG.COL_INICIO_FECHAS).setFontSize(9);
  }

  // Nota de texto libre
  const notaFila = filaLeyenda + leyendaData.length + 1;
  hoja.getRange(notaFila, 1, 1, 2).merge();
  hoja.getRange(notaFila, 1).setValue('💡 Tip:');
  hoja.getRange(notaFila, 1).setFontSize(9);
  hoja.getRange(notaFila, 1).setFontColor(COLORES.gris_oscuro);
  hoja.getRange(notaFila, 1).setFontStyle('italic');
  hoja.getRange(notaFila, CONFIG.COL_INICIO_FECHAS, 1, 6).merge();
  hoja.getRange(notaFila, CONFIG.COL_INICIO_FECHAS).setValue('También puedes escribir comentarios u observaciones libres en cualquier celda');
  hoja.getRange(notaFila, CONFIG.COL_INICIO_FECHAS).setFontSize(9);
  hoja.getRange(notaFila, CONFIG.COL_INICIO_FECHAS).setFontColor(COLORES.gris_oscuro);
  hoja.getRange(notaFila, CONFIG.COL_INICIO_FECHAS).setFontStyle('italic');

  // ================================================
  // ZOOM + ÚLTIMO TOQUE
  // ================================================
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '✅ ¡ROSTER MISAGI creado exitosamente!\n\n' +
    '• ' + CONFIG.CONDUCTORES.length + ' conductores registrados\n' +
    '• ' + fechas.length + ' días generados\n' +
    '• Encabezados congelados\n' +
    '• Leyenda de códigos incluida\n\n' +
    'Ahora cierra este mensaje y empieza a llenar datos.\n' +
    'Usa el menú 🚛 MISAGI Tracking para funciones especiales.'
  );
}

// =============================================
// 📋 MENÚ PERSONALIZADO
// =============================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🚛 MISAGI Tracking');

  menu.addItem('🏗️ Crear hoja desde 0 (reiniciar todo)', 'crearHojaDesde0');
  menu.addSeparator();
  menu.addItem('🎨 Aplicar colores a toda la hoja', 'aplicarFormatoCompleto');
  menu.addItem('✅ Validar todos los códigos', 'validarCodigos');
  menu.addSeparator();
  menu.addItem('📊 Generar resumen del mes actual', 'generarResumenMesActual');
  menu.addItem('📊 Generar resumen completo (todos los meses)', 'generarResumenCompleto');
  menu.addSeparator();

  const subTurnos = ui.createMenu('⚡ Llenado rápido');
  subTurnos.addItem('H1 → H2 → H3 (Hudbay 3 días)', 'llenarHudbay');
  subTurnos.addItem('T1 → T2 (Toquepala 2 días)', 'llenarToquepala');
  subTurnos.addItem('C1 → C2 (Cuajone 2 días)', 'llenarCuajone');
  subTurnos.addItem('A1 → A2 → A3 (Hierro 3 días)', 'llenarHierro');
  subTurnos.addItem('D × 7 (Descanso 7 días)', 'llenarDescanso7');
  subTurnos.addItem('P × 5 (Parado 5 días)', 'llenarParado5');
  menu.addSubMenu(subTurnos);

  menu.addSeparator();
  menu.addItem('➕ Agregar conductor', 'agregarConductor');
  menu.addItem('📋 Ver leyenda', 'mostrarLeyenda');
  menu.addItem('ℹ️ Ayuda', 'mostrarAyuda');

  menu.addToUi();
}

// =============================================
// ✏️ AUTO-FORMATO AL EDITAR
// =============================================
function onEdit(e) {
  const hoja = e.source.getActiveSheet();
  if (hoja.getName() !== CONFIG.NOMBRE_HOJA) return;

  const fila = e.range.getRow();
  const col = e.range.getColumn();

  if (fila < CONFIG.FILA_INICIO_DATOS || col < CONFIG.COL_INICIO_FECHAS) return;

  const valor = e.range.getValue().toString().trim();
  const info = buscarCodigo(valor);

  if (info) {
    e.range.setBackground(info.color);
    e.range.setFontColor(info.texto);
    e.range.setHorizontalAlignment('center');
    e.range.setFontSize(9);
    e.range.setFontWeight('bold');
  } else if (valor !== '') {
    // Comentario libre — permitido
    e.range.setBackground('#F1F5F9');
    e.range.setFontColor('#475569');
    e.range.setHorizontalAlignment('center');
    e.range.setFontSize(8);
    e.range.setFontWeight('normal');
    e.range.setFontStyle('italic');
  } else {
    // Vacía — detectar fondo alternado
    const filaIdx = fila - CONFIG.FILA_INICIO_DATOS;
    e.range.setBackground(filaIdx % 2 === 0 ? COLORES.blanco : COLORES.gris_fondo);
    e.range.setFontColor(COLORES.negro_texto);
    e.range.setFontStyle('normal');
    e.range.setFontWeight('normal');
  }
}

function buscarCodigo(valor) {
  if (!valor) return null;
  const v = valor.trim();
  if (CONFIG.CODIGOS[v]) return CONFIG.CODIGOS[v];
  for (const key in CONFIG.CODIGOS) {
    if (key.toUpperCase() === v.toUpperCase()) return CONFIG.CODIGOS[key];
  }
  return null;
}

// =============================================
// ⚡ LLENADO RÁPIDO
// =============================================
function llenarHudbay() { llenarSecuencia(['H1', 'H2', 'H3']); }
function llenarToquepala() { llenarSecuencia(['T1', 'T2']); }
function llenarCuajone() { llenarSecuencia(['C1', 'C2']); }
function llenarHierro() { llenarSecuencia(['A1', 'A2', 'A3']); }
function llenarDescanso7() { llenarSecuencia(['D', 'D', 'D', 'D', 'D', 'D', 'D']); }
function llenarParado5() { llenarSecuencia(['P', 'P', 'P', 'P', 'P']); }

function llenarSecuencia(codigos) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  const celda = hoja.getActiveCell();
  const fila = celda.getRow();
  const colInicio = celda.getColumn();

  if (fila < CONFIG.FILA_INICIO_DATOS || colInicio < CONFIG.COL_INICIO_FECHAS) {
    SpreadsheetApp.getUi().alert('⚠️ Selecciona una celda en el área de datos primero.');
    return;
  }

  for (let i = 0; i < codigos.length; i++) {
    const rango = hoja.getRange(fila, colInicio + i);
    rango.setValue(codigos[i]);
    const info = CONFIG.CODIGOS[codigos[i]];
    if (info) {
      rango.setBackground(info.color);
      rango.setFontColor(info.texto);
      rango.setHorizontalAlignment('center');
      rango.setFontWeight('bold');
      rango.setFontSize(9);
    }
  }

  SpreadsheetApp.getUi().alert('✅ Llenado: ' + codigos.join(' → '));
}

// =============================================
// 🎨 FORMATO COMPLETO
// =============================================
function aplicarFormatoCompleto() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  const ultimaFila = hoja.getLastRow();
  const ultimaCol = hoja.getLastColumn();

  if (ultimaFila < CONFIG.FILA_INICIO_DATOS) return;

  const numFilas = ultimaFila - CONFIG.FILA_INICIO_DATOS + 1;
  const numCols = ultimaCol - CONFIG.COL_INICIO_FECHAS + 1;

  if (numCols <= 0 || numFilas <= 0) return;

  const rango = hoja.getRange(CONFIG.FILA_INICIO_DATOS, CONFIG.COL_INICIO_FECHAS, numFilas, numCols);
  const valores = rango.getValues();
  const fondos = [];
  const textos = [];

  for (let f = 0; f < valores.length; f++) {
    fondos[f] = [];
    textos[f] = [];
    const bgBase = f % 2 === 0 ? COLORES.blanco : COLORES.gris_fondo;

    for (let c = 0; c < valores[f].length; c++) {
      const val = valores[f][c].toString().trim();
      const info = buscarCodigo(val);
      if (info) {
        fondos[f][c] = info.color;
        textos[f][c] = info.texto;
      } else if (val !== '') {
        fondos[f][c] = '#F1F5F9';
        textos[f][c] = '#475569';
      } else {
        fondos[f][c] = bgBase;
        textos[f][c] = COLORES.negro_texto;
      }
    }
  }

  rango.setBackgrounds(fondos);
  rango.setFontColors(textos);
  rango.setHorizontalAlignment('center');
  rango.setFontWeight('bold');
  rango.setFontSize(9);

  SpreadsheetApp.getUi().alert('✅ ¡Colores aplicados!');
}

// =============================================
// ✅ VALIDACIÓN
// =============================================
function validarCodigos() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  const ultimaFila = hoja.getLastRow();
  const ultimaCol = hoja.getLastColumn();

  const rango = hoja.getRange(CONFIG.FILA_INICIO_DATOS, CONFIG.COL_INICIO_FECHAS,
    ultimaFila - CONFIG.FILA_INICIO_DATOS + 1, ultimaCol - CONFIG.COL_INICIO_FECHAS + 1);
  const valores = rango.getValues();
  const nombres = hoja.getRange(CONFIG.FILA_INICIO_DATOS, 1,
    ultimaFila - CONFIG.FILA_INICIO_DATOS + 1, 1).getValues();

  const errores = [];
  for (let f = 0; f < valores.length; f++) {
    for (let c = 0; c < valores[f].length; c++) {
      const val = valores[f][c].toString().trim();
      if (val === '') continue;
      if (!buscarCodigo(val)) {
        errores.push('• ' + (nombres[f][0] || 'Fila ' + (f + CONFIG.FILA_INICIO_DATOS)) +
          ': "' + val + '"');
      }
    }
  }

  if (errores.length === 0) {
    SpreadsheetApp.getUi().alert('✅ ¡Todos los códigos son válidos!');
  } else {
    SpreadsheetApp.getUi().alert('⚠️ ' + errores.length + ' valores no reconocidos (pueden ser comentarios):\n\n' +
      errores.slice(0, 15).join('\n') + (errores.length > 15 ? '\n... y más' : ''));
  }
}

// =============================================
// ➕ AGREGAR CONDUCTOR
// =============================================
function agregarConductor() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('➕ Nuevo Conductor', 'Nombre completo (APELLIDOS NOMBRES):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const nombre = resp.getResponseText().trim().toUpperCase();
  if (!nombre) { ui.alert('⚠️ Ingresa un nombre.'); return; }

  const resp2 = ui.prompt('Cargo', 'Cargo del conductor (ej: Conductor, Operador):', ui.ButtonSet.OK_CANCEL);
  const cargo = resp2.getSelectedButton() === ui.Button.OK ? resp2.getResponseText().trim() : 'Conductor';

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);

  // Buscar última fila con nombre
  const nombres = hoja.getRange(CONFIG.FILA_INICIO_DATOS, 1, hoja.getLastRow(), 1).getValues();
  let ultimaFila = CONFIG.FILA_INICIO_DATOS;
  for (let i = 0; i < nombres.length; i++) {
    if (nombres[i][0].toString().trim() !== '') {
      ultimaFila = CONFIG.FILA_INICIO_DATOS + i;
    }
  }

  const nuevaFila = ultimaFila + 1;
  hoja.getRange(nuevaFila, 1).setValue(nombre);
  hoja.getRange(nuevaFila, 1).setFontWeight('bold').setFontSize(9);
  hoja.getRange(nuevaFila, 2).setValue(cargo);
  hoja.getRange(nuevaFila, 2).setFontSize(8).setFontColor(COLORES.gris_medio);
  hoja.setRowHeight(nuevaFila, 28);

  // Fondo alternado
  const idx = nuevaFila - CONFIG.FILA_INICIO_DATOS;
  const bg = idx % 2 === 0 ? COLORES.blanco : COLORES.gris_fondo;
  hoja.getRange(nuevaFila, 1, 1, hoja.getLastColumn()).setBackground(bg);

  ui.alert('✅ Conductor "' + nombre + '" agregado en fila ' + nuevaFila);
}

// =============================================
// 📊 RESUMEN MENSUAL
// =============================================
function generarResumenMesActual() {
  const hoy = new Date();
  generarResumenMes(hoy.getMonth(), hoy.getFullYear());
}

function generarResumenCompleto() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  const fechasRow = hoja.getRange(4, CONFIG.COL_INICIO_FECHAS, 1, hoja.getLastColumn() - CONFIG.COL_INICIO_FECHAS + 1).getValues()[0];
  const mesesRow = hoja.getRange(3, CONFIG.COL_INICIO_FECHAS, 1, hoja.getLastColumn() - CONFIG.COL_INICIO_FECHAS + 1).getValues()[0];

  const procesados = new Set();
  // Usar la estructura de meses para generar
  const mesesNombres = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
    'JULIO', 'AGOSTO', 'SETIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

  // Buscar meses con datos
  for (let i = 0; i < mesesRow.length; i++) {
    const mesStr = mesesRow[i].toString().trim().toUpperCase();
    const mesIdx = mesesNombres.indexOf(mesStr);
    if (mesIdx >= 0) {
      // Inferir año basado en posición
      const anio = mesIdx >= 3 ? 2025 : 2026; // Abr-Dic 2025, Ene-Mar 2026
      const key = mesIdx + '-' + anio;
      if (!procesados.has(key)) {
        procesados.add(key);
        generarResumenMes(mesIdx, anio);
      }
    }
  }

  SpreadsheetApp.getUi().alert('✅ Resúmenes generados: ' + procesados.size + ' meses');
}

function generarResumenMes(mes, anio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.NOMBRE_HOJA);
  const mesesNombres = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
    'JULIO', 'AGOSTO', 'SETIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

  const nombreHoja = '📊 ' + mesesNombres[mes] + ' ' + anio;
  let hr = ss.getSheetByName(nombreHoja);
  if (hr) ss.deleteSheet(hr);
  hr = ss.insertSheet(nombreHoja);

  // Título
  hr.getRange(1, 1, 1, 10).merge();
  hr.getRange(1, 1).setValue('📊 RESUMEN ' + mesesNombres[mes] + ' ' + anio);
  hr.getRange(1, 1).setBackground(COLORES.azul_navy);
  hr.getRange(1, 1).setFontColor(COLORES.blanco);
  hr.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  hr.setRowHeight(1, 40);

  // Obtener grupos de operaciones
  const grupos = [];
  const gruposSet = new Set();
  for (const info of Object.values(CONFIG.CODIGOS)) {
    if (!['DESCANSO', 'PARADO', 'MEDIA', 'FERIADO'].includes(info.grupo) && !gruposSet.has(info.grupo)) {
      gruposSet.add(info.grupo);
      grupos.push(info.grupo);
    }
  }

  // Encabezados
  const encabezados = ['CONDUCTOR', ...grupos, 'TRABAJADOS', 'DESCANSO', 'PARADO', '% ACTIVIDAD'];
  for (let i = 0; i < encabezados.length; i++) {
    hr.getRange(3, i + 1).setValue(encabezados[i]);
  }
  hr.getRange(3, 1, 1, encabezados.length).setBackground(COLORES.verde_principal);
  hr.getRange(3, 1, 1, encabezados.length).setFontColor(COLORES.blanco);
  hr.getRange(3, 1, 1, encabezados.length).setFontWeight('bold');

  // Datos de conductores
  const nombres = [];
  for (let f = CONFIG.FILA_INICIO_DATOS; f <= hoja.getLastRow(); f++) {
    const n = hoja.getRange(f, 1).getValue().toString().trim();
    if (n) nombres.push({ nombre: n, fila: f });
  }

  // Encontrar columnas del mes
  const totalCols = hoja.getLastColumn();
  const colsMes = [];
  for (let c = CONFIG.COL_INICIO_FECHAS; c <= totalCols; c++) {
    const val = hoja.getRange(4, c).getValue();
    const mesHeader = hoja.getRange(3, c).getValue().toString().trim().toUpperCase();
    // Chequear si pertenece al mes correcto
    if (mesHeader === mesesNombres[mes] || mesHeader === '') {
      // Verificar por fecha si está disponible  
      colsMes.push(c);
    }
  }

  let filaResumen = 4;
  for (const cond of nombres) {
    const conteo = {};
    grupos.forEach(g => conteo[g] = 0);
    let trabajados = 0, descanso = 0, parado = 0;

    for (const c of colsMes) {
      const val = hoja.getRange(cond.fila, c).getValue().toString().trim();
      const info = buscarCodigo(val);
      if (info) {
        if (info.grupo === 'DESCANSO') descanso++;
        else if (info.grupo === 'PARADO') parado++;
        else if (!['MEDIA', 'FERIADO'].includes(info.grupo)) {
          trabajados++;
          if (conteo[info.grupo] !== undefined) conteo[info.grupo]++;
        }
      }
    }

    hr.getRange(filaResumen, 1).setValue(cond.nombre);
    let col = 2;
    for (const g of grupos) {
      hr.getRange(filaResumen, col).setValue(conteo[g] || 0);
      col++;
    }
    hr.getRange(filaResumen, col).setValue(trabajados);
    hr.getRange(filaResumen, col + 1).setValue(descanso);
    hr.getRange(filaResumen, col + 2).setValue(parado);

    const total = trabajados + descanso + parado;
    const pct = total > 0 ? Math.round(trabajados / total * 100) + '%' : '0%';
    hr.getRange(filaResumen, col + 3).setValue(pct);

    // Color alternado
    const bg = (filaResumen % 2 === 0) ? COLORES.blanco : COLORES.gris_fondo;
    hr.getRange(filaResumen, 1, 1, encabezados.length).setBackground(bg);

    filaResumen++;
  }

  // Auto-resize
  for (let c = 1; c <= encabezados.length; c++) {
    hr.autoResizeColumn(c);
  }
}

// =============================================
// 📋 LEYENDA Y AYUDA
// =============================================
function mostrarLeyenda() {
  let msg = '📋 CÓDIGOS DE OPERACIÓN:\n\n';
  const grupos = {};
  for (const [codigo, info] of Object.entries(CONFIG.CODIGOS)) {
    if (!grupos[info.grupo]) grupos[info.grupo] = [];
    if (!grupos[info.grupo].includes(codigo + ' = ' + info.nombre)) {
      grupos[info.grupo].push(codigo + ' = ' + info.nombre);
    }
  }
  for (const [grupo, codigos] of Object.entries(grupos)) {
    msg += '🔹 ' + grupo + ':\n';
    codigos.forEach(c => msg += '   ' + c + '\n');
    msg += '\n';
  }
  msg += '💡 Texto libre = comentario/observación';
  SpreadsheetApp.getUi().alert(msg);
}

function mostrarAyuda() {
  SpreadsheetApp.getUi().alert(
    '🚛 MISAGI TRACKING — AYUDA\n\n' +
    'LLENAR DATOS:\n' +
    '• Escribe códigos (H1, T2, D, P...) → color automático\n' +
    '• Texto libre → comentario (gris itálica)\n\n' +
    'LLENADO RÁPIDO:\n' +
    '• Selecciona celda → Menú → ⚡ Llenado rápido\n\n' +
    'NUEVO CONDUCTOR:\n' +
    '• Menú → ➕ Agregar conductor\n\n' +
    'NUEVA OPERACIÓN:\n' +
    '• Extensiones → Apps Script → editar CODIGOS en CONFIG\n' +
    '• Ejemplo: "B1": {color:"#059669", texto:"#fff", nombre:"Bambas 1", grupo:"BAMBAS"}\n\n' +
    'REPORTES:\n' +
    '• Menú → 📊 Generar resumen mensual'
  );
}
