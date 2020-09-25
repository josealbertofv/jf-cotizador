function ver_partida(){
  var hoja_partidas = buscar_hoja(nombre_hoja_partidas);
  
  var hoja_partidas_esta_activa = verificar_hoja_activa(hoja_partidas);
  
  if( hoja_partidas_esta_activa ){
    
    var hoja_plantilla = buscar_hoja(nombre_hoja_plantilla);
    
    var registro = [[]];
    registro = obtener_registro_por_celda_activa_dimension_variable(hoja_partidas);
    
    crear_plantilla_partida_directo(registro);
    
    activar_hoja(hoja_plantilla);
  }
}

//-----------------------------------------------------------------------------------------------

function eliminar_partida(){
  var obj_hoja_partidas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( nombre_hoja_partidas );
  var obj_hoja_remota = SpreadsheetApp.openById(id_bd_jfe).getSheetByName( cotejar_nombre_hoja_bd_jfe(obj_hoja_partidas) );
  
  if( verificar_hoja_activa(obj_hoja_partidas) ){
    eliminar_sincronizar_registro_activo(obj_hoja_partidas, obj_hoja_remota); 
  }
}

//-----------------------------------------------------------------------------------------------