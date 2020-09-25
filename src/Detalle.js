function crear_detalle(){
  
  var obj_hoja = buscar_hoja( nombre_hoja_detalle );
  var obj_hoja_rango = buscar_hoja( nombre_hoja_bd );
  var fila = 1;
  var cant_filas = 10;
  
  obj_hoja.clear();
  obj_hoja.getRange(1,1,100,100).setDataValidation(null);
  
  fila = insertar_encabezado_detalle(obj_hoja, fila, cant_filas);
  fila = insertar_linea_detalle(obj_hoja, obj_hoja_rango, fila, cant_filas);
  fila = insertar_pie_detalle(obj_hoja, fila, cant_filas);
  
}

//---------------------------------------------------------------------------------------------------

function insertar_encabezado_detalle(obj_hoja, int_fila, int_cant_filas){
  
  var fila = int_fila;
  
  var contenido = []; 
  var estilos_letras = []; 
  var alineaciones_hztal = []; 
  var alineaciones_vtcal = []; 
  var bordes = []; 
  var fondos = []; 
  var validaciones = [];
  
  var formulas = [];
  formulas[0] = '=ROUND(IFERROR(100*M'+ (int_fila+int_cant_filas+2).toString() +'/I'+ (int_fila+int_cant_filas+2).toString() +',0),2)';
  formulas[1] = '=ROUND(sum(J'+ (fila+1).toString() +':L'+ (fila+1).toString() +'),2)';
  
  // Armado del contenido y estilos del encabezado
  
  // Matriz de contenidos
  contenido[0] = ["", "", "", "", "", "Precio", "", "Cantidad", "", "Utilidad (%)", "Imprevistos (%)", "Negociación (%)", "Otro (%)", "Total añadido %", "Total (S/)"];
  contenido[1] = ["Capítulo", "Partida", "Item", "Descripción", "Unidad", "unitario (S/)", "Cantidad", "Redondeada", "Total (S/)", "", "", "", formulas[0], formulas[1], "Cotización"];
  
  // Matriz de estilos de letras (negrillas)
  estilos_letras[0] = ["bold","bold","bold","bold","bold","bold","bold", "bold","bold","bold","bold","bold","bold","bold","bold"];
  estilos_letras[1] = ["bold","bold","bold","bold","bold","bold","bold", "bold","bold","bold","bold","bold","bold","bold","bold"];
  
  // Matriz de alineación horizontal
  alineaciones_hztal[0] = ["center","center","center","center","center","center","center", "center","center","center","center","center","center","center","center"];
  alineaciones_hztal[1] = ["center","center","center","center","center","center","center", "center","center","center","center","center","center","center","center"];
  
  // Matriz de alineación vertical
  alineaciones_vtcal[0] = ["top","top","top","top","top","top","top", "top","top","top","top","top","top","top","top"];
  alineaciones_vtcal[1] = ["top","top","top","top","top","top","top", "top","top","top","top","top","top","top","top"];
  
  // Matriz de bordes
  bordes[0] =[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  bordes[1] =[borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo];
  
  // Matriz de colores de fondo
  fondos[0] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  fondos[1] = [null, null, null, null, null, null, null, null, null, fondo_azul, fondo_azul, fondo_azul, null, null, null];
  
  // Matriz de validaciones
  validaciones[0] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  validaciones[1] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  
  // Inserción de las líneas del encabezado
  var i; 
  for (e in contenido){ 
    i = parseInt(e,10);
    fila = insertar_linea(obj_hoja, i+1, contenido[i], estilos_letras[i], alineaciones_hztal[i], alineaciones_vtcal[i], bordes[i], fondos[i], validaciones[0] );
  }
  
  return fila;
}

//---------------------------------------------------------------------------------------------------

function agregar_1_fila_ctz(){
  agregar_filas_ctz(1);
}

//---------------------------------------------------------------------------------------------------

function agregar_3_fila_ctz(){
  agregar_filas_ctz(3);
}

//---------------------------------------------------------------------------------------------------

function agregar_5_fila_ctz(){
  agregar_filas_ctz(5);
}

//---------------------------------------------------------------------------------------------------

function agregar_filas_ctz(int_cant_filas){
  
  var hoja_detalle = buscar_hoja(nombre_hoja_detalle);
  var hoja_rango_detalle = buscar_hoja(nombre_hoja_bd);
  var hoja_detalle_esta_activa = verificar_hoja_activa(hoja_detalle);
  
  if(hoja_detalle_esta_activa){
    
    var primera_fila = buscar_fila(hoja_detalle, 'Total (S/)', 9);
    var ultima_fila = buscar_fila(hoja_detalle, 'Subtotal (S/)', 8);
    var fila_activa = hoja_detalle.getCurrentCell().getRow();
    
    if(fila_activa > primera_fila && fila_activa < ultima_fila){
      
      hoja_detalle.insertRows(fila_activa, int_cant_filas);
      insertar_linea_detalle(hoja_detalle, hoja_rango_detalle, fila_activa, int_cant_filas);
      
    }else{
      anunciar('La celda seleccionada debe estar entre los items');
    }
  }
}

//---------------------------------------------------------------------------------------------------

function insertar_linea_detalle(obj_hoja, obj_hoja_rango, int_fila, int_cant_filas){
  
  var fila = int_fila;
  
  var contenido = []; 
  var estilos_letras = []; 
  var alineaciones_hztal = []; 
  var alineaciones_vtcal = []; 
  var bordes = []; 
  var fondos = []; 
  var validaciones = []; 
  
  // Insertar ítem
  
  // Armado del contenido y estilos del encabezado
  
  // Matriz de estilos de letras (negrillas)
  estilos_letras[0] = ["normal","normal","normal","normal","normal","normal","normal", "normal","normal","normal","normal","normal","normal","normal","normal"];
  
  // Matriz de alineación horizontal
  alineaciones_hztal[0] = ["center","center","center","left","center","center","center", "center","right","right","right","right","right","right","right"];
  
  // Matriz de alineación vertical
  alineaciones_vtcal[0] = ["top","top","top","top","top","top","top", "top","top","top","top","top","top","top","top"];
  
  // Matriz de bordes
  bordes[0] =[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  
  // Matriz de colores de fondo
  fondos[0] = [fondo_azul, fondo_azul, null, null, null, null, fondo_azul, null, null, null, null, null, null, null, null];
  
  // Se crea la regla de validación
  var valores_rango = obj_hoja_rango.getRange('A1:A');
  var regla = SpreadsheetApp.newDataValidation().requireValueInRange(valores_rango).build();
  
  // Matriz de validaciones
  validaciones[0] = [null, regla, null, null, null, null, null, null, null, null, null, null, null, null, null];
  
  var i = 0;
  for (var c = 1; c <= int_cant_filas; c++){ 
    
    // Fórmulas de las líneas del detalle de cotización
    var capitulo = '';
    var partida = '';
    var item = '';
    var descripcion = '=IFNA(VLOOKUP(B'+fila+','+ nombre_rango_bd_entera +',2,false),"--")';
    var unidad = '=IFNA(VLOOKUP(B'+fila+','+ nombre_rango_bd_entera +',3,false),"--")';
    //var rendimiento = '=IFNA(VLOOKUP(B'+fila+','+ nombre_rango_bd_entera +',4,false),"--")';
    var precio_unit = '=IFNA(VLOOKUP(B'+fila+','+ nombre_rango_bd_entera +',4,false),"--")';
    var cantidad = '';
    var cantidad_redon = '=roundup(G'+fila+')';
    var total_pen = '=IFERROR(F'+fila+'*H'+fila+',"")';
    var utilidad = '=ROUND(I'+fila+'*$J$2/100,2)';
    var imprevistos = '=ROUND(I'+fila+'*$K$2/100,2)';
    var negociacion = '=ROUND(I'+fila+'*$L$2/100,2)';
    var otro = '=ROUND(I'+fila+'*$M$2/100,2)';
    var total_anadido = '=ROUND(sum(J'+fila+':M'+fila+'),2)';
    var total_ctz = '=ROUND(sum(I'+fila+':M'+fila+'),2)';
    
    // Matriz de contenidos
    contenido[0] = [capitulo, partida, item, descripcion, unidad, precio_unit, cantidad, cantidad_redon, total_pen, utilidad, imprevistos, negociacion, otro, total_anadido, total_ctz];
    
    fila = insertar_linea(obj_hoja, fila, contenido[i], estilos_letras[i], alineaciones_hztal[i], alineaciones_vtcal[i], bordes[i], fondos[i], validaciones[i] );
  }
  
  return fila;
}

//---------------------------------------------------------------------------------------------------

function insertar_pie_detalle(obj_hoja, int_fila, int_cant_filas){
  var fila = int_fila;
  
  var contenido = []; 
  var estilos_letras = []; 
  var alineaciones_hztal = []; 
  var alineaciones_vtcal = []; 
  var bordes = []; 
  var fondos = []; 
  var validaciones = []; 
  
  // Matriz de estilos de letras (negrillas)
  estilos_letras[0] = ["normal","normal","normal","normal","normal","normal","normal", "bold","bold","bold","bold","bold","bold","bold","bold"];
  estilos_letras[1] = ["normal","normal","normal","normal","normal","normal","normal", "normal","normal","normal","normal","normal","bold","normal","normal"];
  estilos_letras[2] = ["normal","normal","normal","normal","normal","normal","normal", "normal","normal","normal","normal","normal","normal","bold","bold"];
  estilos_letras[3] = ["normal","normal","normal","normal","normal","normal","normal", "normal","normal","normal","normal","normal","normal","bold","bold"];
  estilos_letras[4] = ["normal","normal","normal","normal","normal","normal","normal", "normal","normal","normal","normal","normal","normal","bold","bold"];
  
  // Matriz de alineación horizontal
  alineaciones_hztal[0] = ["center","center","center","center","center","center","center", "right","right","right","right","right","right","right","right"];
  alineaciones_hztal[1] = ["center","center","center","center","center","center","center", "center","center","center","center","center","center","center","center"];
  alineaciones_hztal[2] = ["center","center","center","center","center","center","center", "center","center","center","center","center","center","right","right"];
  alineaciones_hztal[3] = ["center","center","center","center","center","center","center", "center","center","center","center","center","center","right","right"];
  alineaciones_hztal[4] = ["center","center","center","center","center","center","center", "center","center","center","center","center","center","right","right"];
  
  // Matriz de alineación vertical
  alineaciones_vtcal[0] = ["top","top","top","top","top","top","top", "top","top","top","top","top","top","top","top"];
  alineaciones_vtcal[1] = ["top","top","top","top","top","top","top", "top","top","top","top","top","top","top","top"];
  alineaciones_vtcal[2] = ["top","top","top","top","top","top","top", "top","top","top","top","top","top","top","top"];
  alineaciones_vtcal[3] = ["top","top","top","top","top","top","top", "top","top","top","top","top","top","top","top"];
  alineaciones_vtcal[4] = ["top","top","top","top","top","top","top", "top","top","top","top","top","top","top","top"];
  
  // Matriz de bordes
  bordes[0] =[null, null, null, null, null, null, null, null, borde_arriba, borde_arriba, borde_arriba, borde_arriba, borde_arriba, borde_arriba, borde_arriba];
  bordes[1] =[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  bordes[2] =[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  bordes[3] =[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  bordes[4] =[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  
  // Matriz de colores de fondo
  fondos[0] = [null, null, null, null, null, null, null, null, null, null, null, null, fondo_azul, null, null];
  fondos[1] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  fondos[2] = [null, null, null, null, null, null, null, null, null, null, null, null, null, fondo_amarillo, fondo_amarillo];
  fondos[3] = [null, null, null, null, null, null, null, null, null, null, null, null, null, fondo_amarillo, fondo_amarillo];
  fondos[4] = [null, null, null, null, null, null, null, null, null, null, null, null, null, fondo_amarillo, fondo_amarillo];
  
  // Matriz de validaciones
  validaciones[0] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  validaciones[1] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  validaciones[2] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  validaciones[3] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  validaciones[4] = [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null];
  
  var formulas = [];
  formulas[0] = '=SUM(I'+ (fila - int_cant_filas).toString()+':I'+ (fila -1).toString() +')';
  formulas[1] = '=SUM(K'+ (fila - int_cant_filas).toString()+':J'+ (fila -1).toString() +')';
  formulas[2] = '=SUM(K'+ (fila - int_cant_filas).toString()+':K'+ (fila -1).toString() +')';
  formulas[3] = '=SUM(L'+ (fila - int_cant_filas).toString()+':L'+ (fila -1).toString() +')';
  formulas[4] = '=SUM(N'+ (fila - int_cant_filas).toString()+':N'+ (fila -1).toString() +')';
  formulas[5] = '=SUM(O'+ (fila - int_cant_filas).toString()+':O'+ (fila -1).toString() +')';
  formulas[6] = '=O' + fila.toString();
  formulas[7] = '=ROUND(O' + (fila+2).toString() + '*0.18, 2)';
  formulas[8] = '=ROUND(O' + (fila+2).toString() + '+O' + (fila+3).toString() + ', 2)';
  
  // Matriz de contenido
  contenido[0] = ["", "", "", "", "", "", "", "Subtotal (S/)", formulas[0], formulas[1], formulas[2], formulas[3], "", formulas[4], formulas[5]];
  contenido[1] = ["", "", "", "", "", "", "", "", "", "", "", "", "Otro (S/)", "", ""];
  contenido[2] = ["", "", "", "", "", "", "", "", "", "", "", "", "", "Subtotal", formulas[6]];
  contenido[3] = ["", "", "", "", "", "", "", "", "", "", "", "", "", "IGV (18%)", formulas[7]];
  contenido[4] = ["", "", "", "", "", "", "", "", "", "", "", "", "", "Total", formulas[8]];
  
  var i; 
  for (e in contenido){ 
    i = parseInt(e,10);
    fila = insertar_linea(obj_hoja, fila, contenido[i], estilos_letras[i], alineaciones_hztal[i], alineaciones_vtcal[i], bordes[i], fondos[i], validaciones[i] );
  }
   
  return fila;
}
