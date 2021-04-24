// V-mar.03.2021
// Mi marca desde ubuntu

function onOpen() {
  crearMenu();
}
// Test: 1st changes
//-----------------------------------------------------------------------------------------------

function crearMenu(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Cotizador')
    .addSubMenu(ui.createMenu('Inicio')
      .addItem('Crear nuevo marco de trabajo', 'crear_marco_trabajo')  // Hay que crear una advertencia     
      .addItem('Función de prueba', 'funcion_prueba'))
    .addSeparator()
    .addSubMenu(ui.createMenu('01-Cotización')
      .addItem('Crear documento', 'crear_documento'))
    .addSeparator()
    .addSubMenu(ui.createMenu('02-Detalle')
      .addItem('Añadir 1 fila', 'agregar_1_fila_ctz')
      .addItem('Añadir 3 filas', 'agregar_3_fila_ctz')
      .addItem('Añadir 5 filas', 'agregar_5_fila_ctz')
      .addItem('Crear cotización', 'crear_cotizacion')     
      .addItem('Nueva plantilla de cotización', 'crear_detalle'))
    .addSeparator()
    .addSubMenu(ui.createMenu('03-Pantilla')
      .addItem('Nueva plantilla de partida', 'crear_plantilla_partida')
      .addItem('Agregar fila a la plantilla', 'agregar_fila_plantilla')
      .addItem('Guardar partida', 'guardar_partida'))
    .addToUi();  
  
  ui.createMenu('Base de datos')
    .addSubMenu(ui.createMenu('Base de datos')
      .addItem('Refrescar base de datos', 'refrescar_bd'))
    .addSeparator()
    .addSubMenu(ui.createMenu('04-Partidas')
      .addItem('Ver partida', 'ver_partida')
      .addItem('Eliminar partida', 'eliminar_partida'))
    .addSeparator()
    .addSubMenu(ui.createMenu('05-Mano de obra')
      .addItem('Agregar Mano de Obra', 'agregar_mano_obra')
      .addItem('Eliminar Mano de Obra', 'eliminar_mano_obra'))
    .addSeparator()
    .addSubMenu(ui.createMenu('06-Materiales')
      .addItem('Agregar Material', 'agregar_materiales')
      .addItem('Eliminar Material', 'eliminar_materiales'))
    .addSeparator()
    .addSubMenu(ui.createMenu('07-Equipos')
      .addItem('Agregar Equipo', 'agregar_equipos')
      .addItem('Eliminar Equipo', 'eliminar_equipos'))
    .addSeparator()
    .addSubMenu(ui.createMenu('08-Insumos')
      .addItem('Agregar Suministro', 'agregar_suministros')
      .addItem('Eliminar Suministro', 'eliminar_suministros'))
    .addToUi();  
}

//-----------------------------------------------------------------------------------------------

function crear_marco_trabajo(){
  
  var hoja_01 = crear_hoja( nombre_hoja_cotizacion );
  var hoja_02 = crear_hoja( nombre_hoja_detalle );
  var hoja_03 = crear_hoja( nombre_hoja_plantilla );
  var hoja_04 = crear_hoja( nombre_hoja_partidas );
  var hoja_05 = crear_hoja( nombre_hoja_mano_obra );
  var hoja_06 = crear_hoja( nombre_hoja_materiales );
  var hoja_07 = crear_hoja( nombre_hoja_equipos );
  var hoja_08 = crear_hoja( nombre_hoja_suministros );
  var hoja_09 = crear_hoja( nombre_hoja_bd );
  
  refrescar_bd();
  hoja_09.hideSheet();
  
  crear_plantilla_partida_directo();
  crear_detalle();
  crear_cotizacion();
  
  app_ctz.setActiveSheet( app_ctz.getSheetByName( nombre_hoja_detalle ) );
  anunciar('Ya puede hacer su cotización');
}

//-----------------------------------------------------------------------------------------------
