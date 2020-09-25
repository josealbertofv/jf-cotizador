// De entrada tiene: (1) objeto base de datos, (2) string del nombre de la hoja, (3) objeto hoja en la app, (4) string de nombre del rango a crear

function importar_bd(obj_bd, str_nombre_hoja_bd, obj_hoja, str_nombre_rango){
  
  var hoja_bd = obj_bd.getSheetByName(str_nombre_hoja_bd); // se especifíca la hoja en la BD
  
  // Se obtienen los valores desde la BD
  var ultima_fila = hoja_bd.getLastRow();
  var ultima_columna = hoja_bd.getLastColumn();
  var valores = hoja_bd.getRange(1,1,ultima_fila,ultima_columna).getValues();
  
  // Se copian los valores de la BD y se decora la hoja
  obj_hoja.getRange(1,1,ultima_fila,ultima_columna).setValues(valores);
  obj_hoja.getRange(1,1,1,ultima_columna).setBorder(false, false, true, false, null, null).setFontWeight("bold").setHorizontalAlignment("center");
  obj_hoja.setColumnWidth(2,280);
  
  // Se crea el rango de valores
  if(str_nombre_rango != null){
    var range = obj_hoja.getRange(2, 1, ultima_fila, ultima_columna);
    app_ctz.setNamedRange(str_nombre_rango, range);
  }
}

//-----------------------------------------------------------------------------------------------

function actualizar_hoja_bd(){
  // En caso de añadir alguna otra tabla al sistema, solo hay que modificar este tramo
  var hoja_1 = buscar_hoja(nombre_hoja_partidas);
  var hoja_2 = buscar_hoja(nombre_hoja_mano_obra);
  var hoja_3 = buscar_hoja(nombre_hoja_materiales);
  var hoja_4 = buscar_hoja(nombre_hoja_equipos);
  var hoja_5 = buscar_hoja(nombre_hoja_suministros);
  var hojas = [ hoja_1, hoja_2, hoja_3, hoja_4, hoja_5];
  
  // Se declaran variables a usar
  var hoja_bd = buscar_hoja(nombre_hoja_bd);
  hoja_bd.clear();
  var rango_bd;
  var rango;
  var cant_col_rango_anterior = 0;
  var fila_ini = 1;
  
  for(e in hojas){
    rango = obtener_rango_de_datos( hojas[e] );
    fila_ini = fila_ini + cant_col_rango_anterior;
    rango_bd = hoja_bd.getRange(fila_ini, 1, rango.getNumRows(), rango.getNumColumns() );
    rango.copyTo(rango_bd, {contentsOnly:true});
    cant_col_rango_anterior = rango.getNumRows();
  }
  // Se crea el rango
  app_ctz.setNamedRange(nombre_rango_bd_entera, hoja_bd.getDataRange() );
}

//-----------------------------------------------------------------------------------------------

function crear_rango(obj_hoja, str_nombre_rango){
  var ultima_fila = hoja_bd.getLastRow();
  var ultima_columna = hoja_bd.getLastColumn();
  
  var range = obj_hoja.getRange(2, 1, ultima_fila, ultima_columna);
  app_ctz.setNamedRange(str_nombre_rango, range);
}

//-----------------------------------------------------------------------------------------------

function refrescar_bd(){
  
  var bd_jfelectrico = SpreadsheetApp.openById(id_bd_jfe);
  
  var hoja_04 = buscar_hoja( nombre_hoja_partidas );
  var hoja_05 = buscar_hoja( nombre_hoja_mano_obra );
  var hoja_06 = buscar_hoja( nombre_hoja_materiales );
  var hoja_07 = buscar_hoja( nombre_hoja_equipos );
  var hoja_08 = buscar_hoja( nombre_hoja_suministros );
  
  importar_bd(bd_jfelectrico, 'Partidas', hoja_04, 'partidas');
  importar_bd(bd_jfelectrico, 'Mano de Obra', hoja_05, 'mano_obra');
  importar_bd(bd_jfelectrico, 'Materiales', hoja_06, 'materiales');
  importar_bd(bd_jfelectrico, 'Equipos', hoja_07, 'equipos');
  importar_bd(bd_jfelectrico, 'Suministros', hoja_08, 'suministros');
  actualizar_hoja_bd();
}

//-----------------------------------------------------------------------------------------------