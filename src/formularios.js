function agregar_mano_obra() {
  mostrar_barra_lateral(buscar_hoja(nombre_hoja_mano_obra), 'mano_obra_01');
}

//-----------------------------------------------------------------------------------------------

function agregar_materiales() {
  mostrar_barra_lateral(buscar_hoja(nombre_hoja_materiales), 'materiales_01');
}

//-----------------------------------------------------------------------------------------------

function agregar_equipos() {
  mostrar_barra_lateral(buscar_hoja(nombre_hoja_equipos), 'equipos_01');
}

//-----------------------------------------------------------------------------------------------

function agregar_suministros() {
  mostrar_barra_lateral(buscar_hoja(nombre_hoja_suministros), 'suministros_01');
}

//-----------------------------------------------------------------------------------------------

function agregar_insumo_desde_formulario(form_data){
  
  var registro = [[]];
  var hoja = app_ctz.getActiveSheet();
  var hoja_remota = SpreadsheetApp.openById(id_bd_jfe).getSheetByName( cotejar_nombre_hoja_bd_jfe(hoja) );
  
  registro = obtener_datos_desde_formulario(hoja, form_data);
  
  guardar_sincronizar_registro(hoja, registro, hoja_remota);
}

//-----------------------------------------------------------------------------------------------

function obtener_datos_desde_formulario(obj_hoja, form_data){
  var registro = [[]]; 
  var nombre_hoja = obj_hoja.getName();
  
  var usuario = Session.getActiveUser().getEmail();
  var costo_usd_fecha = '=ROUND( index( GOOGLEFINANCE("currency:USDPEN", "price", INDIRECT( ADDRESS(ROW(), COLUMN()-1, 4) ) ) ,2,2) , 2) ';
  var costo_usd_hoy = '=ROUND(GOOGLEFINANCE("currency:USDPEN"),2)';
  
  if(nombre_hoja == nombre_hoja_materiales){
    registro[0][0] = form_data.codigo;
    registro[0][1] = form_data.descripcion;
    registro[0][2] = form_data.unidad;
    registro[0][3] = form_data.costo_pen;
    registro[0][4] = form_data.marca;
    registro[0][5] = form_data.proveedor;
    registro[0][6] = form_data.url;
    registro[0][7] = form_data.peso_kg;
    registro[0][8] = form_data.volumen;
    registro[0][9] = form_data.hoja_tecnica;
    registro[0][10] = usuario;
    registro[0][11] = form_data.fecha;  
    registro[0][12] = costo_usd_fecha;
    registro[0][13] = costo_usd_hoy;
    
  }else if(nombre_hoja == nombre_hoja_equipos){
    registro[0][0] = form_data.codigo;
    registro[0][1] = form_data.descripcion;
    registro[0][2] = form_data.unidad;
    registro[0][3] = form_data.costo_pen;
    registro[0][4] = form_data.marca;
    registro[0][5] = form_data.proveedor;
    registro[0][6] = form_data.url;
    registro[0][7] = form_data.peso_kg;
    registro[0][8] = form_data.volumen;
    registro[0][9] = form_data.hoja_tecnica;
    registro[0][10] = usuario;
    registro[0][11] = form_data.fecha;  
    registro[0][12] = costo_usd_fecha;
    registro[0][13] = costo_usd_hoy;
    
  }else if(nombre_hoja == nombre_hoja_suministros){
    registro[0][0] = form_data.codigo;
    registro[0][1] = form_data.descripcion;
    registro[0][2] = form_data.unidad;
    registro[0][3] = form_data.costo_pen;
    registro[0][4] = form_data.marca;
    registro[0][5] = form_data.proveedor;
    registro[0][6] = form_data.url;
    registro[0][7] = form_data.peso_kg;
    registro[0][8] = form_data.volumen;
    registro[0][9] = form_data.hoja_tecnica;
    registro[0][10] = usuario;
    registro[0][11] = form_data.fecha;  
    registro[0][12] = costo_usd_fecha;
    registro[0][13] = costo_usd_hoy;
    
  }else if(nombre_hoja == nombre_hoja_mano_obra){
    registro[0][0] = form_data.codigo;
    registro[0][1] = form_data.descripcion;
    registro[0][2] = form_data.unidad;
    registro[0][3] = form_data.costo_pen;
    registro[0][4] = form_data.basico;
    registro[0][5] = form_data.gratificacion;
    registro[0][6] = form_data.vacaciones;
    registro[0][7] = form_data.transporte;
    registro[0][8] = form_data.alimentacion;
    registro[0][9] = form_data.nocturno;
    registro[0][10] = form_data.bonificacion;
    registro[0][11] = usuario;
    registro[0][12] = form_data.fecha;  
    registro[0][13] = costo_usd_fecha;
    registro[0][14] = costo_usd_hoy;
  }
  
  return registro;
}

//-----------------------------------------------------------------------------------------------

function eliminar_mano_obra(){
  var obj_hoja_mano_obra = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( nombre_hoja_mano_obra );
  var obj_hoja_remota = SpreadsheetApp.openById(id_bd_jfe).getSheetByName( cotejar_nombre_hoja_bd_jfe(obj_hoja_mano_obra) );
  
  if( verificar_hoja_activa(obj_hoja_mano_obra) ){
    eliminar_sincronizar_registro_activo(obj_hoja_mano_obra, obj_hoja_remota); 
  }
}

//-----------------------------------------------------------------------------------------------

function eliminar_materiales(){
  var obj_hoja_materiales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( nombre_hoja_materiales );
  var obj_hoja_remota = SpreadsheetApp.openById(id_bd_jfe).getSheetByName( cotejar_nombre_hoja_bd_jfe(obj_hoja_materiales) );
  
  if( verificar_hoja_activa(obj_hoja_materiales) ){
    eliminar_sincronizar_registro_activo(obj_hoja_materiales, obj_hoja_remota); 
  }
}

//-----------------------------------------------------------------------------------------------

function eliminar_equipos(){
  var obj_hoja_equipos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( nombre_hoja_equipos );
  var obj_hoja_remota = SpreadsheetApp.openById(id_bd_jfe).getSheetByName( cotejar_nombre_hoja_bd_jfe(obj_hoja_equipos) );
  
  if( verificar_hoja_activa(obj_hoja_equipos) ){
    eliminar_sincronizar_registro_activo(obj_hoja_equipos, obj_hoja_remota); 
  }
}

//-----------------------------------------------------------------------------------------------

function eliminar_suministros(){
  var obj_hoja_suministros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( nombre_hoja_suministros );
  var obj_hoja_remota = SpreadsheetApp.openById(id_bd_jfe).getSheetByName( cotejar_nombre_hoja_bd_jfe(obj_hoja_suministros) );
  
  if( verificar_hoja_activa(obj_hoja_suministros) ){
    eliminar_sincronizar_registro_activo(obj_hoja_suministros, obj_hoja_remota); 
  }
}