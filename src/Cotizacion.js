function crear_cotizacion(){
  
  var obj_hoja = buscar_hoja( nombre_hoja_cotizacion );
  var obj_hoja_detalle = buscar_hoja( nombre_hoja_detalle );
  var fila = 1;
  
  obj_hoja.clear();
  obj_hoja.getRange(1,1,10,100).setDataValidation(null);
  
  fila = insertar_datos_cliente(obj_hoja, obj_hoja_detalle, fila);
  fila = insertar_encabezado_cotizacion(obj_hoja, obj_hoja_detalle, fila);
  var fila_pie = fila;
  fila = insertar_contenido_cotizacion(obj_hoja, obj_hoja_detalle, fila);
  fila = insertar_pie_cotizacion(obj_hoja, obj_hoja_detalle, fila_pie);
  
  app_ctz.setActiveSheet(obj_hoja);
}

//---------------------------------------------------------------------------------------------------

function insertar_datos_cliente(obj_hoja, obj_hoja_detalle, int_fila){
  var fila = int_fila;
  var contenido = [];
  var estilos_letras = [];
  var alineaciones_hztal = [];
  var alineaciones_vtcal = [];
  var bordes = [];
  var fondos = [];
  var validaciones = [];
  
  contenido[0] = [null, null, "Datos del Cliente", null, null, null];
  contenido[1] = [null, "Nombre cliente: ", null, null, null, null];
  contenido[2] = [null, "Nombre corto de cliente: ", null, null, null, null];
  contenido[3] = [null, "Tipo de ID: ", null, null, null, null];
  contenido[4] = [null, "ID: ", null, null, null, null];
  contenido[5] = [null, "Nombre atención: ", null, null, null, null];
  contenido[6] = [null, "Correo-e: ", null, null, null, null];
  contenido[7] = [null, "Tlf: ", null, null, null, null];
  contenido[8] = [null, "Nro cotización: ", null, null, null, null];
  contenido[9] = [null, "Descripción servicio: ", null, null, null, null];
  contenido[10] = [null, "Fecha: ", null, null, null, null];
  contenido[11] = [null, "Cant. días de entrega: ", null, null, null, null];
  contenido[12] = [null, "%  de 1er pago: ", null, null, null, null];
  contenido[13] = [null, "%  de 2do pago: ", null, null, null, null];
  
  
  estilos_letras[0] = [null, null, 'bold', null, null, null];
  estilos_letras[1] = [null, null, null, null, null, null];
  estilos_letras[2] = [null, null, null, null, null, null];
  estilos_letras[3] = [null, null, null, null, null, null];
  estilos_letras[4] = [null, null, null, null, null, null];
  estilos_letras[5] = [null, null, null, null, null, null];
  estilos_letras[6] = [null, null, null, null, null, null];
  estilos_letras[7] = [null, null, null, null, null, null];
  estilos_letras[8] = [null, null, null, null, null, null];
  estilos_letras[9] = [null, null, null, null, null, null];
  estilos_letras[10] = [null, null, null, null, null, null];
  estilos_letras[11] = [null, null, null, null, null, null];
  estilos_letras[12] = [null, null, null, null, null, null];
  estilos_letras[13] = [null, null, null, null, null, null];  
  
  
  alineaciones_hztal[0] = [null, null, 'center', null, null, null];
  alineaciones_hztal[1] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[2] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[3] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[4] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[5] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[6] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[7] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[8] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[9] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[10] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[11] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[12] = [null, 'right', 'left', null, null, null];
  alineaciones_hztal[13] = [null, 'right', 'left', null, null, null];
  
  
  alineaciones_vtcal[0] = [null, null, null, null, null, null];
  alineaciones_vtcal[1] = [null, null, null, null, null, null];
  alineaciones_vtcal[2] = [null, null, null, null, null, null];
  alineaciones_vtcal[3] = [null, null, null, null, null, null];
  alineaciones_vtcal[4] = [null, null, null, null, null, null];
  alineaciones_vtcal[5] = [null, null, null, null, null, null];
  alineaciones_vtcal[6] = [null, null, null, null, null, null];
  alineaciones_vtcal[7] = [null, null, null, null, null, null];
  alineaciones_vtcal[8] = [null, null, null, null, null, null];
  alineaciones_vtcal[9] = [null, null, null, null, null, null];
  alineaciones_vtcal[10] = [null, null, null, null, null, null];
  alineaciones_vtcal[11] = [null, null, null, null, null, null];
  alineaciones_vtcal[12] = [null, null, null, null, null, null];
  alineaciones_vtcal[13] = [null, null, null, null, null, null];
  
  
  bordes[0] = [null, null, null, null, null, null];
  bordes[1] = [null, null, null, null, null, null];
  bordes[2] = [null, null, null, null, null, null];
  bordes[3] = [null, null, null, null, null, null];
  bordes[4] = [null, null, null, null, null, null];
  bordes[5] = [null, null, null, null, null, null];
  bordes[6] = [null, null, null, null, null, null];
  bordes[7] = [null, null, null, null, null, null];
  bordes[8] = [null, null, null, null, null, null];
  bordes[9] = [null, null, null, null, null, null];
  bordes[10] = [null, null, null, null, null, null];
  bordes[11] = [null, null, null, null, null, null];
  bordes[12] = [null, null, null, null, null, null];
  bordes[13] = [null, null, null, null, null, null];
  
  
  fondos[0] = [null, null, null, null, null, null];
  fondos[1] = [null, null, fondo_azul, null, null, null];
  fondos[2] = [null, null, fondo_azul, null, null, null];
  fondos[3] = [null, null, fondo_azul, null, null, null];
  fondos[4] = [null, null, fondo_azul, null, null, null];
  fondos[5] = [null, null, fondo_azul, null, null, null];
  fondos[6] = [null, null, fondo_azul, null, null, null];
  fondos[7] = [null, null, fondo_azul, null, null, null];
  fondos[8] = [null, null, fondo_azul, null, null, null];
  fondos[9] = [null, null, fondo_azul, null, null, null];
  fondos[10] = [null, null, fondo_azul, null, null, null];
  fondos[11] = [null, null, fondo_azul, null, null, null];
  fondos[12] = [null, null, fondo_azul, null, null, null];
  fondos[13] = [null, null, fondo_azul, null, null, null];
  
  
  validaciones[0] = [null, null, null, null, null, null];
  validaciones[1] = [null, null, null, null, null, null];
  validaciones[2] = [null, null, null, null, null, null];
  validaciones[3] = [null, null, null, null, null, null];
  validaciones[4] = [null, null, null, null, null, null];
  validaciones[5] = [null, null, null, null, null, null];
  validaciones[6] = [null, null, null, null, null, null];
  validaciones[7] = [null, null, null, null, null, null];
  validaciones[8] = [null, null, null, null, null, null];
  validaciones[9] = [null, null, null, null, null, null];
  validaciones[10] = [null, null, null, null, null, null];
  validaciones[11] = [null, null, null, null, null, null];
  validaciones[12] = [null, null, null, null, null, null];
  validaciones[13] = [null, null, null, null, null, null];
  
  
  obj_hoja.setColumnWidth(3, 280);
  
  var i = 0;
  for(e in contenido){
    i = parseInt(e, 10);
    fila = insertar_linea(obj_hoja, fila, contenido[i], estilos_letras[i], alineaciones_hztal[i], alineaciones_vtcal[i], bordes[i], fondos[i], validaciones[i] );
  }
  
  return fila;
}

//---------------------------------------------------------------------------------------------------

function insertar_encabezado_cotizacion(obj_hoja, obj_hoja_detalle, int_fila){
  
  var fila = int_fila;
  var contenido = [];
  var estilos_letras = [];
  var alineaciones_hztal = [];
  var alineaciones_vtcal = [];
  var bordes = [];
  var fondos = [];
  var validaciones = [];
  
  contenido[0] = [null, null, null, null, null, null];
  contenido[1] = [null, null, "Cuadro de Costos", null, null, null];
  contenido[2] = [null, null, null, null, null, null];
  contenido[3] = [null, null, null, null, null, null];
  contenido[4] = ["Item", "Código", "Descripción", "Unidad", "Cantidad", "Total"];
  
  estilos_letras[0] = [null, null, null, null, null, null];
  estilos_letras[1] = [null, null, 'bold', null, null, null];
  estilos_letras[2] = [null, null, null, null, null, null];
  estilos_letras[3] = [null, null, null, null, null, null];
  estilos_letras[4] = ['bold', 'bold', 'bold', 'bold', 'bold', 'bold'];
  
  alineaciones_hztal[0] = [null, null, null, null, null, null];
  alineaciones_hztal[1] = [null, null, 'center', null, null, null];
  alineaciones_hztal[2] = [null, null, null, null, null, null];
  alineaciones_hztal[3] = [null, null, null, null, null, null];
  alineaciones_hztal[4] = ['center', 'center', 'center', 'center', 'center', 'center'];
  
  alineaciones_vtcal[0] = [null, null, null, null, null, null];
  alineaciones_vtcal[1] = [null, null, null, null, null, null];
  alineaciones_vtcal[2] = [null, null, null, null, null, null];
  alineaciones_vtcal[3] = [null, null, null, null, null, null];
  alineaciones_vtcal[4] = [null, null, null, null, null, null];
  
  bordes[0] = [null, null, null, null, null, null];
  bordes[1] = [null, null, null, null, null, null];
  bordes[2] = [null, null, null, null, null, null];
  bordes[3] = [borde_arriba_medio, borde_arriba_medio, borde_arriba_medio, borde_arriba_medio, borde_arriba_medio, borde_arriba_medio];
  bordes[4] = [borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo, borde_bajo];
  
  fondos[0] = [null, null, null, null, null, null];
  fondos[1] = [null, null, null, null, null, null];
  fondos[2] = [null, null, null, null, null, null];
  fondos[3] = [null, null, null, null, null, null];
  fondos[4] = [null, null, null, null, null, null];
  
  validaciones[0] = [null, null, null, null, null, null];
  validaciones[1] = [null, null, null, null, null, null];
  validaciones[2] = [null, null, null, null, null, null];
  validaciones[3] = [null, null, null, null, null, null];
  validaciones[4] = [null, null, null, null, null, null];
  
  var i = 0;
  for(e in contenido){
    i = parseInt(e, 10);
    fila = insertar_linea(obj_hoja, fila, contenido[i], estilos_letras[i], alineaciones_hztal[i], alineaciones_vtcal[i], bordes[i], fondos[i], validaciones[i] );
  }
  
  return fila;
}

//---------------------------------------------------------------------------------------------------

function insertar_contenido_cotizacion(obj_hoja, obj_hoja_detalle, int_fila){
  
  var ultima_fila = obj_hoja_detalle.getLastRow();
  var ultima_columna = obj_hoja_detalle.getLastColumn();
  var rango_cotizacion = obj_hoja_detalle.getRange(3, 1, ultima_fila, ultima_columna);
  app_ctz.setNamedRange('cotizacion', rango_cotizacion);
  var formula_1 = '={ ArrayFormula(query(cotizacion,'+'\"select C, B, D, E, H, O where G > 0 \",1))}';
  
  obj_hoja.getRange(int_fila,1,1,1).setValue(formula_1);
  
  return int_fila++;
}

//---------------------------------------------------------------------------------------------------

function insertar_pie_cotizacion(obj_hoja, obj_hoja_detalle, int_fila){

  var fila_ini = int_fila;  
  var fila_fin = obj_hoja.getLastRow();
  obj_hoja.getRange(fila_ini,6,fila_fin-fila_ini+1,1).setHorizontalAlignment('right');
  obj_hoja.getRange(fila_ini,4,fila_fin-fila_ini+1,2).setHorizontalAlignment('center');
  
  var fila = fila_fin+1;
  var contenido = [];
  var estilos_letras = [];
  var alineaciones_hztal = [];
  var alineaciones_vtcal = [];
  var bordes = [];
  var fondos = [];
  var validaciones = [];
  
  var formula_1 = '=SUM(F'+ fila_ini.toString() +':F'+ fila_fin.toString() +')';
  var formula_2 = '=ROUND(F'+ (fila_fin+1).toString() +'*0.18,2)';
  var formula_3 = '=ROUND(F'+ (fila_fin+1).toString() +'+F'+ (fila_fin+2).toString() +', 2)';
  
  contenido[0] = [null, null, null, null, 'Subtotal', formula_1];
  contenido[1] = [null, null, null, null, 'IGV (18%)', formula_2];
  contenido[2] = [null, null, null, null, 'Total', formula_3];
  
  
  estilos_letras[0] = [null, null, null, null, 'bold', 'bold'];
  estilos_letras[1] = [null, null, null, null, 'bold', 'bold'];
  estilos_letras[2] = [null, null, null, null, 'bold', 'bold'];
  
  
  alineaciones_hztal[0] = [null, null, null, null, 'right', 'right'];
  alineaciones_hztal[1] = [null, null, null, null, 'right', 'right'];
  alineaciones_hztal[2] = [null, null, null, null, 'right', 'right'];
  
  
  alineaciones_vtcal[0] = [null, null, null, null, null, null];
  alineaciones_vtcal[1] = [null, null, null, null, null, null];
  alineaciones_vtcal[2] = [null, null, null, null, null, null];
  
  
  bordes[0] = [null, null, null, null, borde_arriba, borde_arriba];
  bordes[1] = [null, null, null, null, null, null];
  bordes[2] = [borde_bajo_medio, borde_bajo_medio, borde_bajo_medio, borde_bajo_medio, borde_bajo_medio, borde_bajo_medio];
  
  
  fondos[0] = [null, null, null, null, null, null];
  fondos[1] = [null, null, null, null, null, null];
  fondos[2] = [null, null, null, null, null, null];
  
  
  validaciones[0] = [null, null, null, null, null, null];
  validaciones[1] = [null, null, null, null, null, null];
  validaciones[2] = [null, null, null, null, null, null];
  
  
  var i = 0;
  for(e in contenido){
    i = parseInt(e, 10);
    fila = insertar_linea(obj_hoja, fila, contenido[i], estilos_letras[i], alineaciones_hztal[i], alineaciones_vtcal[i], bordes[i], fondos[i], validaciones[i] );
  }
  
  return fila;
}


