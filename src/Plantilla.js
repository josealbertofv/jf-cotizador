function crear_plantilla_partida(){
  var hoja_plantilla = buscar_hoja(nombre_hoja_plantilla);
  var hoja_plantilla_esta_activa = verificar_hoja_activa(hoja_plantilla);
  if(hoja_plantilla_esta_activa){
    crear_plantilla_partida_directo();
  }
}

//---------------------------------------------------------------------------------------------------

function crear_plantilla_partida_directo(arr_registro){
  
  var hoja_03 = buscar_hoja(nombre_hoja_plantilla);
  var hoja_04 = buscar_hoja(nombre_hoja_partidas);
  var hoja_05 = buscar_hoja(nombre_hoja_mano_obra);
  var hoja_06 = buscar_hoja(nombre_hoja_materiales);
  var hoja_07 = buscar_hoja(nombre_hoja_equipos);
  var hoja_08 = buscar_hoja(nombre_hoja_suministros);
  
  hoja_03.clear();
  hoja_03.getRange(1,1,100,10).setDataValidation(null);
  
  generar_plantilla_partida(hoja_03, hoja_05, hoja_06, hoja_07, 5, arr_registro)
}

//---------------------------------------------------------------------------------------------------

function agregar_fila_plantilla(){
  var int_fila_partida = 1;
  var hoja_plantilla = buscar_hoja(nombre_hoja_plantilla);
  var esta_activa_hoja_plantilla = verificar_hoja_activa(hoja_plantilla);
  
  if(esta_activa_hoja_plantilla){
    var fila;
    var delimitador = ['Total Mano de obra (S/)', 'Total Materiales (S/)', 'Total Equipos (S/)'];
    var nombres_rango = ['mano_obra', 'materiales', 'equipos'];
    
    var contenidos = [];
    var estilos_letras = [];
    var alineaciones_hztal = [];
    var alineaciones_vtcal = [];
    var bordes = []; 
    var fondos = [];
    var validaciones = [];
    
    var formulas = [];
    // Se delcaran variables para actualizar la fórmula del total de subclase
    var celda_total_subclase;
    var formula;
    var indice_1;
    var indice_2;
    var indice_3;
    var fila_ini;
    var fila_fin;
    
    // Ingreso de contenidos, validaciones y estilos
    //contenidos[0] = ["", str_titulo, "", "", "", "", "", ""];
    estilos_letras[0] = ["normal","normal","normal","normal","normal","normal","normal","normal"];
    alineaciones_hztal[0] = ["center","left","center","center","center","center","right","right"];
    alineaciones_vtcal[0] = ["top","top","top","top","top","top","top","top"];
    bordes[0] =[null, null, null, null, null, null, null, null];
    fondos[0] = [null, null, null, null, null, null, null, null];
    validaciones[0] = [null, null, null, null, null, null, null, null];
    
    var i;
    for(e in delimitador){
      i = parseInt(e, 10);
      fila = buscar_fila(hoja_plantilla, delimitador[i], 6);
      
      formulas[0] = '=IFNA(VLOOKUP(A'+ fila.toString() + ','+ nombres_rango[i] +',2,false),"--")';
      formulas[1] = '=IFNA(VLOOKUP(A'+ fila.toString() +','+ nombres_rango[i] +',3,false),"--")';
      if(nombres_rango[i]=='mano_obra'){ formulas[2] = '=iferror(trunc(E'+ fila.toString() + '*$D$'+(int_fila_partida+1).toString()+'/1,2),"--")';}else{ formulas[2] = '--'; } 
      formulas[3] = '=IFNA(VLOOKUP(A'+ fila.toString() +','+ nombres_rango[i] +',4,false),0)';
      formulas[4] = '=E'+ fila.toString() +'*F'+ fila.toString();
      formulas[5] = '=iferror(trunc(100*(G'+ fila.toString() +'/$G$'+(int_fila_partida+1).toString()+'),2)&"%","--")';
      
      contenidos[0] = ["", formulas[0], formulas[1], formulas[2], "", formulas[3], formulas[4], formulas[5]];
      
      hoja_plantilla.insertRowAfter(fila-1);
      //hoja_plantilla.insertRows(fila, 1);
      insertar_linea(hoja_plantilla, fila, contenidos[0], estilos_letras[0], alineaciones_hztal[0], alineaciones_vtcal[0], bordes[0], fondos[0], validaciones[0] );
      
      // De aquí en adelante se actualiza el subtotal de la subclase
      celda_total_subclase = hoja_plantilla.getRange(fila+1,7);
      formula = celda_total_subclase.getFormula();
      indice_1 = formula.indexOf("(");
      indice_2 = formula.indexOf(":");
      indice_3 = formula.indexOf(")");
      if(indice_2 == -1){
        fila_ini = parseInt( formula.substring(indice_1+2, indice_3), 10);
        fila_fin = fila_ini + 1;
      }else{
        fila_ini = parseInt( formula.substring(indice_1+2, indice_2), 10);
        fila_fin = parseInt( formula.substring(indice_2+2, indice_3), 10) + 1;
      }
      formula = '=SUM(G' + fila_ini.toString() + ':G' + fila_fin.toString() + ')';
      celda_total_subclase.setValue(formula);
    }
  }
}

//---------------------------------------------------------------------------------------------------
// hay que sacar int_cant_filas
function generar_plantilla_partida(obj_hoja_plantilla, obj_hoja_mano_obra, obj_hoja_materiales, obj_hoja_equipos, int_cant_filas, arr_registro){
  var fila_inicio = 1;
  var cant_filas = [5,5,5];
  var filas = [0,0];
  
  filas = insertar_encabezado(obj_hoja_plantilla, arr_registro, fila_inicio, cant_filas);
  
  filas = insertar_encabezado_clasificador(obj_hoja_plantilla, 'Mano de obra', filas[0]);
  filas = insertar_item_partida(obj_hoja_plantilla, 'Mano de obra', obj_hoja_mano_obra,'mano_obra', filas[1], filas[0], fila_inicio, arr_registro);
  cant_filas[0] = filas[1];
  filas = insertar_pie_clasificador(obj_hoja_plantilla, 'Mano de obra', filas[1], filas[0], fila_inicio);
  
  filas = insertar_encabezado_clasificador(obj_hoja_plantilla, 'Materiales', filas[0]);
  filas = insertar_item_partida(obj_hoja_plantilla, 'Materiales', obj_hoja_materiales,'materiales', filas[1], filas[0], fila_inicio, arr_registro);
  cant_filas[1] = filas[1];
  filas = insertar_pie_clasificador(obj_hoja_plantilla, 'Materiales', filas[1], filas[0], fila_inicio);
  
  filas = insertar_encabezado_clasificador(obj_hoja_plantilla, 'Equipos', filas[0]);
  filas = insertar_item_partida(obj_hoja_plantilla, 'Equipos', obj_hoja_equipos,'equipos', filas[1], filas[0], fila_inicio, arr_registro);
  cant_filas[2] = filas[1];
  filas = insertar_pie_clasificador(obj_hoja_plantilla, 'Equipos', filas[1], filas[0], fila_inicio);
  
  filas = insertar_encabezado(obj_hoja_plantilla, arr_registro, fila_inicio, cant_filas);
}

//---------------------------------------------------------------------------------------------------

function insertar_encabezado(obj_hoja, arr_partida, int_fila_inicio, int_cant_filas){
  var partida = arr_partida;
  if(arr_partida == null){
    partida = ["","","","","","","",""]; 
  }else{
    partida = [arr_partida[0], arr_partida[1], arr_partida[2], arr_partida[4],"","","",""];
  }
  
  var fila_1 = ( int_fila_inicio+4+1+int_cant_filas[0] ).toString();
  var fila_2 = ( int_fila_inicio+4+3+int_cant_filas[0]+int_cant_filas[1] ).toString();
  var fila_3 = ( int_fila_inicio+4+5+int_cant_filas[0]+int_cant_filas[1]+int_cant_filas[2] ).toString();
  
  var formula_total ="= G"+ fila_1 +"+G"+ fila_2 +"+G"+ fila_3 ;
    
  var contenidos = [];
  var estilos_letras = [];
  var alineaciones_hztal = [];
  var alineaciones_vtcal = [];
  var bordes = []; 
  var fondos = [];
  var validaciones = [];
  
  // Matriz de contenidos
  contenidos[0] = ["Partida", "Descripción", "Unidad", "Un / día", "", "", "Costo / Un (S/)", ""];
  contenidos[1] = [partida[0], partida[1] , partida[2] , partida[3] ,"","", formula_total,""];
  contenidos[2] = ["", "", "", "", "", "", "", ""];
  contenidos[3] = ["Código", "Descripción", "Unidad", "Cuadrilla", "Cantidad", "P. unit", "Precio", "%"];
   
  // Matriz de estilos de letras (negrillas)
  estilos_letras[0] = ["bold","bold","bold","bold","normal","normal","bold","normal"];
  estilos_letras[1] = ["normal","normal","normal","normal","normal","normal","normal","normal"];
  estilos_letras[2] = ["normal","normal","normal","normal","normal","normal","normal","normal"];
  estilos_letras[3] = ["bold","bold","bold","bold","bold","bold","bold","bold"];
  
  // Matriz de alineación horizontal
  alineaciones_hztal[0] = ["center","center","center","center","left","right","left","left"];
  alineaciones_hztal[1] = ["center","left","center","center","center","left","right","left"];
  alineaciones_hztal[2] = ["left","left","left","left","left","left","left","left"];
  alineaciones_hztal[3] = ["center","center","center","center","center","right","right","right"];
  
  // Matriz de alineación vertical
  alineaciones_vtcal[0] = ["top","top","top","top","top","top","top","top"];
  alineaciones_vtcal[1] = ["top","top","top","top","top","top","top","top"];
  alineaciones_vtcal[2] = ["top","top","top","top","top","top","top","top"];
  alineaciones_vtcal[3] = ["top","top","top","top","top","top","top","top"];
  
  // Matriz de bordes
  bordes[0] =[borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo];
  bordes[1] =[null, null, null, null, null, null, null, null];
  bordes[2] =[null, null, null, null, null, null, null, null];
  bordes[3] =[borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo];
  
  // Matriz de colores de fondo
  fondos[0] = [null, null, null, null, null, null, null, null];
  fondos[1] = [fondo_azul, fondo_azul, fondo_azul, fondo_azul, null, null, null, null];
  fondos[2] = [null, null, null, null, null, null, null, null];
  fondos[3] = [null, null, null, null, null, null, null, null];
  
  // Matriz de validaciones
  validaciones[0] = [null, null, null, null, null, null, null, null];
  validaciones[1] = [null, null, null, null, null, null, null, null];
  validaciones[2] = [null, null, null, null, null, null, null, null];
  validaciones[3] = [null, null, null, null, null, null, null, null];

  // Inserción de contenidos y estilos
  var i;
  for(e in validaciones){ 
    i = parseInt(e, 10);
    insertar_linea(obj_hoja, i+1, contenidos[i], estilos_letras[i], alineaciones_hztal[i], alineaciones_vtcal[i], bordes[i], fondos[i], validaciones[i] );
  }
  
  // Configuración de anchos de columnas
  obj_hoja.setColumnWidth(1,100).setColumnWidth(2,280).setColumnWidth(3,60).setColumnWidth(4,60).setColumnWidth(5,60).setColumnWidth(6,60).setColumnWidth(7,70).setColumnWidth(8,70);
  
  //hoja.getRange(ini+1,2,ini+c,1).setWrap(true); Para que se muestre todo el contenido autodimensionando la celda
  
  var cant_fil = contenidos.length;
  
  return [cant_fil + int_fila_inicio, 0];
}

//---------------------------------------------------------------------------------------------------

function insertar_encabezado_clasificador(obj_hoja, str_titulo, int_fila_inicio){
  
  var fila = int_fila_inicio;
  var contenidos = [];
  var estilos_letras = [];
  var alineaciones_hztal = [];
  var alineaciones_vtcal = [];
  var bordes = []; 
  var fondos = [];
  var validaciones = [];
  
  // Ingreso de contenidos, validaciones y estilos
  contenidos[0] = ["", str_titulo, "", "", "", "", "", ""];
  estilos_letras[0] = ["normal","normal","normal","normal","normal","normal","normal","normal"];
  alineaciones_hztal[0] = ["left","left","left","left","left","left","left","left"];
  alineaciones_vtcal[0] = ["top","top","top","top","top","top","top","top"];
  bordes[0] =[null, null, null, null, null, null, null, null];
  fondos[0] = [null, null, null, null, null, null, null, null];
  validaciones[0] = [null, null, null, null, null, null, null, null];
  
  fila = insertar_linea(obj_hoja, fila, contenidos[0], estilos_letras[0], alineaciones_hztal[0], alineaciones_vtcal[0], bordes[0], fondos[0], validaciones[0] );
  
  return [fila, 0];
}

//---------------------------------------------------------------------------------------------------

function insertar_item_partida(obj_hoja, str_titulo, obj_hoja_rango, str_rango, int_cant_filas, int_fila_inicio, int_fila_partida, arr_partida){
    
  var fila = int_fila_inicio;
  var cant_filas = 5;
  var contenidos = [];
  var estilos_letras = [];
  var alineaciones_hztal = [];
  var alineaciones_vtcal = [];
  var bordes = []; 
  var fondos = [];
  var validaciones = [];
  
  var codigos_partida = [];
  var cantidades_partida = [];
  
  var hay_partida = false;
  if(arr_partida != null){ 
    hay_partida = true;
  }
  
  if(hay_partida){
    
    var iniciales_rango = cotejar_nombre_rango_iniciales(str_rango);
    
    var dimension = arr_partida.length - 8;
    var e = 0;
    for(var i = 1; i <= dimension; i++){
      if(arr_partida[8+i-1].toString().substring(0,2) == iniciales_rango){
        codigos_partida[e] = arr_partida[8+i-1];
        cantidades_partida[e] = arr_partida[8+i-1+1];
        e++;
      }
    }
    cant_filas = codigos_partida.length;
  }
  
  // Se crea la regla de validación
  var regla = crear_regla_rango(obj_hoja_rango, 'A2:A');
  
  // Ingreso de contenidos, validaciones y estilos
  estilos_letras[0] = ["normal","normal","normal","normal","normal","normal","normal","normal"];
  alineaciones_hztal[0] = ["center","left","center","center","center","center","right","right"];
  alineaciones_vtcal[0] = ["top","top","top","top","top","top","top","top"];
  bordes[0] =[null, null, null, null, null, null, null, null];
  fondos[0] = [fondo_azul, null, null, null, fondo_azul, null, null, null];
  validaciones[0] = [regla, null, null, null, null, null, null, null];
  
  // Se crea al menos una fila vacía, en caso de que no haya items
  if(cant_filas<1){
    cant_filas = 1;
  }
    
  // A continaución se insertan las líneas de la plantilla y previamente se construyen las fórmulas del contenido
  var formulas = [];
  for(var i = 1; i <= cant_filas; i++){
    
    formulas[0] = '=IFNA(VLOOKUP(A'+ fila.toString() + ','+ str_rango +',2,false),"--")';
    formulas[1] = '=IFNA(VLOOKUP(A'+ fila.toString() +','+ str_rango +',3,false),"--")';
    if(str_rango=='mano_obra'){ formulas[2] = '=iferror(trunc(E'+ fila.toString() + '*$D$'+(int_fila_partida+1).toString()+'/1,2),"--")';}else{ formulas[2] = '--'; } 
    formulas[3] = '=IFNA(VLOOKUP(A'+ fila.toString() +','+ str_rango +',4,false),0)';
    formulas[4] = '=E'+ fila.toString() +'*F'+ fila.toString();
    formulas[5] = '=iferror(trunc(100*(G'+ fila.toString() +'/$G$'+(int_fila_partida+1).toString()+'),2)&"%","--")';
    
    // Se llena el contenido de la línea con las fórmulas
    if(hay_partida){
      contenidos[0] = [codigos_partida[i-1], formulas[0], formulas[1], formulas[2], cantidades_partida[i-1], formulas[3], formulas[4], formulas[5]];
    }else{
      contenidos[0] = ["", formulas[0], formulas[1], formulas[2], "", formulas[3], formulas[4], formulas[5]];
    }
    
    fila = insertar_linea(obj_hoja, fila, contenidos[0], estilos_letras[0], alineaciones_hztal[0], alineaciones_vtcal[0], bordes[0], fondos[0],  validaciones[0]);
  }
  return [fila, cant_filas];
}

//---------------------------------------------------------------------------------------------------

function insertar_pie_clasificador(obj_hoja, str_titulo, int_cant_filas, int_fila_inicio, int_fila_partida){
  var formulas = [];
  var contenidos = [];
  
  var estilos_letras = [];
  var alineaciones_hztal = [];
  var alineaciones_vtcal = [];
  var bordes = []; 
  var fondos = [];
  var validaciones = [];
  
  var fila = int_fila_inicio;
  
  formulas[0] = '=SUM(G' + (fila-int_cant_filas).toString() + ':G' + (fila-1).toString() + ')';
  formulas[1] = '=iferror(trunc(100*(G'+ fila.toString() +'/$G$'+(int_fila_partida+1).toString()+'),2)&"%","--")';
  
  // Ingreso de contenidos, validaciones y estilos
  contenidos[0] = ["", "", "", "", "", "Total " + str_titulo + " (S/)", formulas[0], formulas[1]];
  estilos_letras[0] = ["normal","normal","normal","normal","normal","normal","normal","normal"];
  alineaciones_hztal[0] = ["left","left","left","left","left","right","right","right"];
  alineaciones_vtcal[0] = ["top","top","top","top","top","top","top","top"];
  bordes[0] =[null, null, null, null, null, null, borde_arriba, borde_arriba];
  fondos[0] = [null, null, null, null, null, null, null, null];
  validaciones[0] = [null, null, null, null, null, null, null, null];
    
  fila = insertar_linea(obj_hoja, fila, contenidos[0], estilos_letras[0], alineaciones_hztal[0], alineaciones_vtcal[0], bordes[0], fondos[0], validaciones[0] );
  return [fila, 0];
}

//---------------------------------------------------------------------------------------------------

function guardar_partida(){
  // Se obtiene la app y la bd
  var app_ctz = SpreadsheetApp.getActiveSpreadsheet();
  var bd_jfe = SpreadsheetApp.openById(id_bd_jfe);
  
  // Se obtiene la hoja activa, de trabajo y remota
  var hoja_plantilla = app_ctz.getSheetByName(nombre_hoja_plantilla);
  var hoja_partidas = app_ctz.getSheetByName(nombre_hoja_partidas); 
  var hoja_remoto = bd_jfe.getSheetByName( cotejar_nombre_hoja_bd_jfe(hoja_partidas) );
  
  // Obtención de datos últiles
  var usuario = Session.getActiveUser().getEmail();
  var fecha = new Date();
  var fila = 0;
  var registrar = true;
  var partida=[[]];
    
  // Se asegura que esté en la hoja correspondiente  
  if( verificar_hoja_activa(hoja_plantilla) ){ 
    // Lee los valores primarios desde la hoja Plantilla
    fila = fila + 2;
    partida[0][0] = hoja_plantilla.getRange(fila,1).getValue();
    partida[0][1] = hoja_plantilla.getRange(fila,2).getValue();
    partida[0][2] = hoja_plantilla.getRange(fila,3).getValue();
    partida[0][3] = hoja_plantilla.getRange(fila,7).getValue();
    partida[0][4] = hoja_plantilla.getRange(fila,4).getValue();
    partida[0][5] = usuario;
    partida[0][6] = fecha;
    partida[0][7] = '=ROUND( index( GOOGLEFINANCE("currency:USDPEN", "price", INDIRECT( ADDRESS(ROW(), COLUMN()-1, 4) ) ) ,2,2) , 2) ';
    partida[0][8] = '=ROUND(GOOGLEFINANCE("currency:USDPEN"),2)';
        
    // Se verifica que ningún dato del encabezado esté vacío
    if(partida[0][0] !== "" && partida[0][1] !== "" && partida[0][2] !== "" && partida[0][3] !== "" && partida[0][4] !== ""){
      fila = fila + 4;
      var fila_fin = hoja_plantilla.getLastRow(); // Indica el límite de la búsqueda de items de la partida a guardar
      var valor;
      var ind_campo = 9; // Desde este campo en el registro partidas[[]] se comienzan a almacenar los items de la partida
      for(fila; fila < fila_fin; fila++){
        // Se toma el valor en la columna 1 para descartar los vacíos
        valor = hoja_plantilla.getRange(fila,1,1,1).getValue();
        if(valor !== ""){
          partida[0][ind_campo] = valor;
          partida[0][ind_campo+1] = hoja_plantilla.getRange(fila,5,1,1).getValue();
          // Se verifica que un tiem no tenga cantidad vacía
          if(partida[0][ind_campo+1]==''){
            registrar = false; 
            var insumo_observacion = partida[0][ind_campo];
          }
          ind_campo = ind_campo +2;
        }
      }
      
      if(registrar==true){
        guardar_sincronizar_registro(hoja_partidas, partida, hoja_remoto); // Si todo está bien, se procede a guardar el registro
        
      }else{
        anunciar('Revise la cantidad del insumo ' + insumo_observacion + '. Ninguna cantidad puede estar vacía');
      }
    }else{
      anunciar('Ninguno de los campos del encabezado puede estar vacío');
    }
  }
}

