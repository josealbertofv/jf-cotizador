/*  Lista de funciones

insertar_linea(...)

buscar_hoja(...)

crear_hoja(...)

buscar_fila(...)

encontrar_texto_en_columna(...)

guardar_sincronizar_registro(...)

consultar_continuacion(...)

anunciar(...)

insertar_registro(...)

buscar_registro_por_clave(...)

comparar_registros_por_dimension(...)

comparar_registros_por_campo(...)

comparar_registros_identicos(...)

ordenar_tabla(...)

eliminar_sincronizar_registro_activo(...)

obtener_registro_por_celda_activa_dimension_variable(...)

obtener_registro_por_celda_activa_dimension_estatica(...)

cotejar_nombre_hoja_bd_jfe(...)

verificar_hoja_activa(...)

crear_regla_rango(...)

transponer_matriz(...)

obtener_rango_de_datos(...)

funcion_prueba(...)

*/

function insertar_linea(obj_hoja, int_nro_fila, arr_valores, arr_estilos_letras, arr_alineaciones_hztal, arr_alineaciones_vtcal, arr_bordes, arr_fondos, arr_validaciones ){

  var rango = obj_hoja.getRange(int_nro_fila, 1, 1, arr_valores.length);
  var valores = Array.of(arr_valores);
  var estilos_letras = Array.of(arr_estilos_letras);
  var alineaciones_hztal = Array.of(arr_alineaciones_hztal);
  var alineaciones_vtcal = Array.of(arr_alineaciones_vtcal);
  var validaciones = Array.of(arr_validaciones);
  
  rango.setValues(valores).setFontWeights(estilos_letras).setHorizontalAlignments(alineaciones_hztal).setVerticalAlignments(alineaciones_vtcal);
  
  var borde = [];
  var fondo = [];
  var validacion = [];
  var j = 1;
  for(e in arr_bordes){
    borde = arr_bordes[e];
    if(borde){ 
        obj_hoja.getRange(int_nro_fila, j).setBorder(borde.top, borde.left, borde.bottom, borde.right, borde.vertical, borde.horizontal, borde.color, borde.style);
    }
    
    fondo = arr_fondos[e];
    if(fondo){
      obj_hoja.getRange(int_nro_fila, j).setBackground(fondo);
    }
    
    validacion = arr_validaciones[e];
    if(validacion){
      obj_hoja.getRange(int_nro_fila, j).setDataValidation(validacion);
    }
    j++;
  } 
  return int_nro_fila + 1;
}

//-----------------------------------------------------------------------------------------------

function buscar_hoja(str_nombre_hoja){
  var hoja = null;
  
  if(str_nombre_hoja == null){
    hoja = crear_hoja(str_nombre_hoja);
  }else{  
    hoja = app_ctz.getSheetByName(str_nombre_hoja);
  }
  return hoja;
}

//-----------------------------------------------------------------------------------------------

function crear_hoja(str_nombre){
  
  var nombre = str_nombre;
    
  if(nombre==null){ // Se valida si el argumento es nulo
    nombre = 'Nueva hoja';
  }
    
  var nueva_hoja = app_ctz.getSheetByName(nombre);
  if (nueva_hoja != null) { // Se valida si la hoja ya existe
    app_ctz.deleteSheet(nueva_hoja); // Si la hoja ya existe, se borra
  }
    
  nueva_hoja = app_ctz.insertSheet(); // Se crea la nueva hoja
  nueva_hoja.setName(nombre); // Se nombra la nueva hoja
  
  return nueva_hoja;
} 

//-----------------------------------------------------------------------------------------------

function buscar_fila(obj_hoja, str_palabras, int_columna){
  
  var ultima_fila = obj_hoja.getLastRow(); 
  // Se especifica en qué columna se hará la búsqueda y hasta qué fila
  var rango_busqueda = obj_hoja.getRange(1,int_columna,ultima_fila,1).getValues();
  // Se busca la palabra que coincide y se devuelve la fila donde ocurre
  var i = 0;
  var fila = -1;
  if(ultima_fila<1001 && int_columna<101){  // Límtes de protección
    for(e in rango_busqueda){
      i = parseInt(e, 10);    
      if(rango_busqueda[i]==str_palabras){
        fila = i+1; 
      } 
    }
  }
  return fila;
}

//-----------------------------------------------------------------------------------------------

function encontrar_texto_en_columna(obj_hoja, str_palabras, int_columna){
  
  var ultima_fila = obj_hoja.getLastRow(); 
  // Se especifica en qué columna se hará la búsqueda y hasta qué fila
  var rango_busqueda = obj_hoja.getRange(1,int_columna,ultima_fila,1).getValues();
  // Se busca la palabra que coincide y se devuelve la fila donde ocurre
  var i = 0;
  var fila = -1;
  if(ultima_fila<1001){  // Límte de protección
    for(e in rango_busqueda){
      i = parseInt(e, 10);    
      if(rango_busqueda[i]==str_palabras){
        fila = i+1; 
      } 
    }
  }
  
  var existencia;
  
  if(fila == -1){
    existencia = false;
  }else{
    existencia = true;
  }
  
  return existencia;
}


//-----------------------------------------------------------------------------------------------

function guardar_sincronizar_registro(obj_hoja, arr_registro, obj_hoja_remota){
  
  // Se revisa si hay duplicado de código o descripción
  var hay_duplicado_en_columna_1 = encontrar_texto_en_columna(obj_hoja, arr_registro[0][0], 1);
  var hay_duplicado_en_columna_2 = encontrar_texto_en_columna(obj_hoja, arr_registro[0][1], 2);
  var exito_bd_local = false;
  var exito_bd_remota = false;
  
  // Si hay dupicado, se consulta si aún quiere continuar
  if(hay_duplicado_en_columna_1){
    var respuesta = consultar_continuacion('ALERTA: El registro ' + arr_registro[0][0] + ' ya existe. ¿Está seguro que quiere reemplazarlo? ');
    if(respuesta == 'YES'){
      // En caso afirmativo se guarda el registro, pero primero se elimina el registro existente
      exito_bd_local = eliminar_registro(obj_hoja, arr_registro);
      exito_bd_remota = eliminar_registro(obj_hoja_remota, arr_registro);
      if(exito_bd_local && exito_bd_remota){
        exito_bd_local = insertar_registro(obj_hoja, arr_registro);
        exito_bd_remota = insertar_registro(obj_hoja_remota, arr_registro);
      }
    }
  // Si la descripción está duplicada, no se continúa  se consulta si aún quiere continuar
  }else if(hay_duplicado_en_columna_2){
    var respuesta = consultar_continuacion('ALERTA: Ya existe otro registro con la misma descripción. ¿Quiere añadir otro registro con descripción repetida?');
    if(respuesta == 'YES'){
      // En caso afirmativo se guarda el registro
      exito_bd_local = insertar_registro(obj_hoja, arr_registro);
      exito_bd_remota = insertar_registro(obj_hoja_remota, arr_registro);
    }
  // Si ninguno de los dos campos está duplicado, se guarda el registro
  }else{
    exito_bd_local = insertar_registro(obj_hoja, arr_registro);
    exito_bd_remota = insertar_registro(obj_hoja_remota, arr_registro);
  }
  
  // Anuncia el éxito o el fracaso de guardar la base de datos
  if(exito_bd_local && exito_bd_remota){
    anunciar('El registro ' + arr_registro[0][0] + ' se guardó con éxito');
  }else{
    if(exito_bd_local == false){
      anunciar('ERROR: No se puede guardar en la tabla local');
    }else{
      anunciar('ERROR: No se puede guardar en la tabla remota');
    }
  }
}

//-----------------------------------------------------------------------------------------------

function insertar_registro(obj_hoja, arr_registro){
  // Se obtienen la última fila y última columna
  var ultima_fila = obj_hoja.getLastRow();
  var ultima_columna = obj_hoja.getLastColumn();
  
  // Se insertan los valores
  obj_hoja.getRange(ultima_fila+1, 1, 1, arr_registro[0].length).setValues(arr_registro);
  
  // Se procede a verificar que el registro esté en la hoja
  var confirmacion = false;
  
  var registro_en_tabla = buscar_registro_por_clave(obj_hoja, arr_registro);
  
  var tienen_dimensiones_iguales = comparar_registros_por_dimension(arr_registro, registro_en_tabla); 
  
  var los_dos_primeros_valores_son_iguales = comparar_registros_por_campo(arr_registro, registro_en_tabla, 1) && comparar_registros_por_campo(arr_registro, registro_en_tabla, 2);
  
  if(tienen_dimensiones_iguales && los_dos_primeros_valores_son_iguales){
    confirmacion = true;
  }
  
  return confirmacion;
}

//-----------------------------------------------------------------------------------------------


function eliminar_sincronizar_registro_activo(obj_hoja, obj_hoja_remota){
  
  var exito = false;
  var exito_bd_local = false;
  var exito_bd_remota = false;
  
  // Se busca la fila activa
  var hoja_activa = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var rango_activo = hoja_activa.getActiveRange();
  
  var valor_en_rango_activo = rango_activo.getValue();
  
  if(valor_en_rango_activo != ''){
    var fila = rango_activo.getRow();
  
    // Se obtiene el registro
    var ultima_columna = hoja_activa.getLastColumn();
    var arr_registro = hoja_activa.getRange(fila,1,1,ultima_columna).getValues();
  
    // Se procede a eliminar el registro en ambas bases de datos
    var respuesta = consultar_continuacion('ALERTA: ¿Está usted seguro que quiere eliminar el registro ' + arr_registro[0][0] + '? No podrá revertir esta operación');
    if(respuesta == 'YES'){
      exito_bd_local = eliminar_registro(obj_hoja, arr_registro);
      exito_bd_remota = eliminar_registro(obj_hoja_remota, arr_registro);
    }
  
    // Se enuncia el éxito o el error en la eliminación
    if(exito_bd_local && exito_bd_remota){
      exito = true;
      anunciar('Se eliminó el registro ' + arr_registro[0][0] + ' con éxito');
    }else{
      anunciar('ERROR: No se puso eliminar el registro ' + arr_registro[0][0] + ' correctamente');
    }   
  }else{
    anunciar('La celda seleccionada está vacía');
  }
  
  return exito;
}

//-----------------------------------------------------------------------------------------------

function eliminar_registro(obj_hoja, arr_registro){
  var exito = false;
  var fila;
  
  var existe_el_registro = encontrar_texto_en_columna(obj_hoja, arr_registro[0][0], 1);
  
  if(existe_el_registro){
    fila = buscar_fila(obj_hoja, arr_registro[0][0], 1);
    obj_hoja.deleteRow(fila);
    
    existe_el_registro = encontrar_texto_en_columna(obj_hoja, arr_registro[0][0], 1);
    if(existe_el_registro == false){
      exito = true;
    }
  }
  return exito;
}

//-----------------------------------------------------------------------------------------------

function consultar_continuacion(str_mensaje) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     str_mensaje,
     'No podrá revertir esta decisión',
    ui.ButtonSet.YES_NO);
  return result;
}

//-----------------------------------------------------------------------------------------------

function anunciar(str_mensaje){
  var ui = SpreadsheetApp.getUi();
  ui.alert(str_mensaje);
}

//-----------------------------------------------------------------------------------------------

function buscar_registro_por_clave(obj_hoja, arr_registro){
  // Se busca la fila y luego se toma el registro entero desde la fila
  var fila = buscar_fila(obj_hoja, arr_registro[0][0], 1);
  var registro = obj_hoja.getRange(fila, 1, 1, arr_registro[0].length).getValues();
  return registro;
}

//-----------------------------------------------------------------------------------------------

function comparar_registros_con_excepcion(arr_registro_1, arr_registro_2, arr_excepcion){
  var son_iguales = true;
  var no_es_columna_excepcional = true;
  
  for(var i = 0; i < arr_registro_1[0].length; i++){
    no_es_columna_excepcional = (arr_excepcion.indexOf(i+1) == -1);
    // En cada item del registro verifica que sean iguales y que no sea columna excepción
    if( arr_registro_1[0][i] != arr_registro_2[0][i] && no_es_columna_excepcional ){
      son_iguales = false;
    }
  }
  return son_iguales;
}

//-----------------------------------------------------------------------------------------------

function comparar_registros_por_dimension(arr_registro_1, arr_registro_2){
  var igual_dimension = false;
  if( arr_registro_1[0].length == arr_registro_2[0].length ){
    igual_dimension = true; 
  }
  return igual_dimension;  
}

//-----------------------------------------------------------------------------------------------

function comparar_registros_por_campo(arr_registro_1, arr_registro_2, int_columna){
  var mismo_valor = false;
  if( arr_registro_1[0][int_columna-1] == arr_registro_2[0][int_columna-1] ){
    mismo_valor = true;
  }
  return mismo_valor;
}

//-----------------------------------------------------------------------------------------------

function comparar_registros_identicos(arr_registro_1, arr_registro_2){
  var son_identicos = true;
  
  son_identicos = comparar_registros_por_dimension(arr_registro_1, arr_registro_2);
  
  for(var i = 0; i < arr_registro_1[0].length; i++){
    if(arr_registro_1[0][i] != arr_registro_2[0][i]){
      son_identicos = false;
    }
  }
  return son_identicos;
}

//-----------------------------------------------------------------------------------------------

function ordenar_tabla(obj_hoja, int_columna){
  // Se obtienen los límites de la hoja
  var ultima_fila = obj_hoja.getLastRow();
  var ultima_columna = obj_hoja.getLastColumn();
  // Se ordena la hoja, si hay un registro, por lo menos
  if(ultima_fila>2){ 
    obj_hoja.getRange(2, 1, ultima_fila-1, ultima_columna).sort(int_columna);
  }
}

//-----------------------------------------------------------------------------------------------

function obtener_registro_por_celda_activa_dimension_variable(obj_hoja){
  
  var valor_celda_activa = obj_hoja.getActiveRange().getValue();
    
  if(valor_celda_activa != ""){
    var fila = obj_hoja.getCurrentCell().getRow();
    var registro = [[]];      
    var valor_siguiente = "temp";
    var i = 0;
    // Se llena el registro con el valor de celda por celda
    while(valor_siguiente != "" && i<100){
      registro[i] = obj_hoja.getRange(fila, i+1,1,1).getValue();
      i++;
      valor_siguiente = obj_hoja.getRange(fila, i+1,1,1).getValue();
    }
  }else{
    SpreadsheetApp.getUi().alert('El registro está vacío');
  }
  return registro;
}

//-----------------------------------------------------------------------------------------------

function obtener_registro_por_celda_activa_dimension_estatica(obj_hoja){

  var valor_celda_activa = obj_hoja.getActiveRange().getValue();
    
  if(valor_celda_activa != ""){
    var fila = obj_hoja.getCurrentCell().getRow();
    var cant_columnas = obj_hoja.getLastColumn();
    var registro = [[]];      
    var valor_siguiente = "temp";
    // Se guarda el registro
    registro = obj_hoja.getRange(fila, 1, 1, cant_columnas).getValues();
  }else{
    SpreadsheetApp.getUi().alert('El registro está vacío');
  }
  return registro;
}

//-----------------------------------------------------------------------------------------------

function cotejar_nombre_hoja_bd_jfe(obj_hoja){
  
  var nombre_hoja = obj_hoja.getName();
  var nombre_hoja_bd_jfe;
  
  if(nombre_hoja == nombre_hoja_partidas ){
    nombre_hoja_bd_jfe = nombre_hoja_bd_jfe_partidas;
  }else if(nombre_hoja == nombre_hoja_mano_obra ){
    nombre_hoja_bd_jfe = nombre_hoja_bd_jfe_mano_obra;
  }else if(nombre_hoja == nombre_hoja_materiales ){
    nombre_hoja_bd_jfe = nombre_hoja_bd_jfe_materiales;
  }else if(nombre_hoja == nombre_hoja_equipos ){
    nombre_hoja_bd_jfe = nombre_hoja_bd_jfe_equipos;
  }else if(nombre_hoja == nombre_hoja_suministros ){
    nombre_hoja_bd_jfe = nombre_hoja_bd_jfe_suministros;
  }else{
    anunciar('No se puede conectar con la base de datos');
  }
  
  return nombre_hoja_bd_jfe;
}

//-----------------------------------------------------------------------------------------------

function cotejar_nombre_rango_iniciales(str_rango){
  
  var iniciales_rango;
  
  if(str_rango == nombre_rango_mano_obra ){
    iniciales_rango = 'MO';
  }else if(str_rango == nombre_rango_materiales ){
    iniciales_rango = 'MA';
  }else if(str_rango == nombre_rango_equipos ){
    iniciales_rango = 'EQ';
  }else if(str_rango == nombre_rango_suministros ){
    iniciales_rango = 'SU';
  }else{
    anunciar('No se puede conectar con la base de datos');
  }
  
  return iniciales_rango;
}

//-----------------------------------------------------------------------------------------------

function verificar_hoja_activa(obj_hoja){
  // Se obtiene App y la hoja activa
  var app_ctz = SpreadsheetApp.getActiveSpreadsheet();
  var hoja_activa = app_ctz.getActiveSheet();
  
  var es_hoja_activa = false;
  
  if(obj_hoja.getName() == hoja_activa.getName()){
    es_hoja_activa = true;
  }else{
    anunciar('Para ejecutar esta acción debe estar en la hoja ' + obj_hoja.getName());
  }
  
  return es_hoja_activa;
}

//-----------------------------------------------------------------------------------------------

function crear_regla_rango(obj_hoja, str_rango){
    var valores_rango = obj_hoja.getRange(str_rango);
    var regla = SpreadsheetApp.newDataValidation().requireValueInRange(valores_rango).build();
    return regla;
  }

//-----------------------------------------------------------------------------------------------

function activar_hoja(obj_hoja){
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(obj_hoja);
}

//-----------------------------------------------------------------------------------------------

function transponer_matriz(array) {
  return array[0].map((col, i) => array.map(row => row[i]));
}

//-----------------------------------------------------------------------------------------------

function obtener_rango_de_datos(obj_hoja){
  
  var cant_filas_1 = obj_hoja.getLastRow();
  var cant_columnas_1 = obj_hoja.getLastColumn();
  
  var rango = obj_hoja.getRange(2,1,cant_filas_1-1,cant_columnas_1);
  
  return rango;
}

//-----------------------------------------------------------------------------------------------

function mostrar_barra_lateral(obj_hoja, archivo_html){
  app_ctz.setActiveSheet(obj_hoja);
  
  var html = HtmlService.createTemplateFromFile(archivo_html).evaluate().setTitle("Formulario de " + obj_hoja.getName() );
  SpreadsheetApp.getUi().showSidebar(html);
  
  // El siguiente modo es para caja de diálogo
  /*var html = HtmlService.createHtmlOutputFromFile(archivo_html);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Item');*/
}

//-----------------------------------------------------------------------------------------------

// Esta función es genérica para hacer cualquier prueba
function funcion_prueba(){
  
  
  
  //Logger.log('El valor de ' + ' ' + ' es: ' + ' ' );
  SpreadsheetApp.getUi().alert('Se ejecutó la función de prueba hasta este paso');
}
