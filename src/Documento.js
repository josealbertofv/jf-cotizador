function crear_documento() {
  
  var obj_hoja = buscar_hoja('01-Cotización');  // TEMPORAL
  
  // Se obtiene los datos del cliente, que estarán como contenido del encabezado y pies de página
  var fila = 1;
  var col = 2;
  fila = buscar_fila(obj_hoja, 'Nombre cliente: ', col); 
  var nombre_cliente = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'Nombre corto de cliente: ', col); 
  var nombre_resumido_cliente = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'Tipo de ID: ', col);
  var tipo_id = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'ID: ', col);
  var id_cliente = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'Nombre atención: ', col);
  var nombre_atencion = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'Correo-e: ', col);
  var correo_cliente = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'Tlf: ', col);
  var tlf_cliente = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'Nro cotización: ', col);
  var nro_ctz = obj_hoja.getRange(fila, col+1,1,1).getValue(); 
  
  fila = buscar_fila(obj_hoja, 'Descripción servicio: ', col);
  var desc_servicio = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'Fecha: ', col);
  var fecha = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, 'Cant. días de entrega: ', col);
  var dias_entrega = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, '%  de 1er pago: ', col);
  var primer_pago = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  fila = buscar_fila(obj_hoja, '%  de 2do pago: ', col);
  var segundo_pago = obj_hoja.getRange(fila, col+1,1,1).getValue();
  
  // Una vez que se tienen los datos, comienza el proceso de ubicación y creación del documento
  
  // Se especifica desde qué hoja de cálculo se trabaja
  var hoja_calculo_id = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // A continuación no se debería especificar la hoja, sin embargo toma la primera hoja
  
  // A continuación se localiza la carpeta donde está la hoja de cálculo activa
  var folder = DriveApp.getFileById(hoja_calculo_id).getParents().next();
  
  // Se crea el nombre que tendrá el documento
  var nombre_doc = 'ctz ' + nombre_resumido_cliente + ' ' + nro_ctz;
  
  // Se genera una copia del documento plantilla
  var doc_ctz = DriveApp.getFileById(plantilla_doc_id).makeCopy(nombre_doc, folder);
  
  // Se obtiene el ID del nuevo documento
  var doc_ctz_id = doc_ctz.getId();
  
  // Se rellena el encabezado de la primera página y luego de las siguientes
  var header_1 = DocumentApp.openById(doc_ctz_id).getHeader().getParent().getChild(2);
  header_1.replaceText('##nro_ctz##', nro_ctz);
  header_1.replaceText('##fecha##', fecha);
  var header_2 = DocumentApp.openById(doc_ctz_id).getHeader();
  header_2.replaceText('##nro_ctz##', nro_ctz);
  header_2.replaceText('##fecha##', fecha);
  
  // Se obtiene el cuerpo del documento creado
  var body = DocumentApp.openById(doc_ctz_id).getBody();
  
  // Se rellena el contenido general en el cuerpo
  body.replaceText('##desc_servicio##', desc_servicio);
  body.replaceText('##nombre_cliente##', nombre_cliente);
  body.replaceText('##tipo_id##', tipo_id);
  body.replaceText('##id_cliente##', id_cliente);
  body.replaceText('##nombre_atencion##', nombre_atencion);
  
  // Se rellena el contenido en las condiciones
  body.replaceText('##dias_entrega##', dias_entrega);
  body.replaceText('##primer_pago##', primer_pago);
  body.replaceText('##segundo_pago##', segundo_pago);
  
  // Ahora se procede con el CUADRO DE COTIZACIÓN
  // Desde la hoja de cálculo:
  // Se cuenta desde donde comienza el contenido de la cotización, en la hoja de cálculo
  var titulo_cuadro_costos = 'Cuadro de Costos'; // Se puede cambiar
  
  var inicio_ctz = buscar_fila(obj_hoja, titulo_cuadro_costos, 3);
  inicio_ctz = inicio_ctz + 3;

  // Se especifica el rango de valores que se copia desde la hoja de cálculo
  //var fila_fin = findInColumn(SpreadsheetApp.getActiveSpreadsheet(), 'E', 'Subtotal')+2;
  var fila_fin = buscar_fila(obj_hoja, 'Subtotal', 5)+2;
  
  // Se toma el contenido que irá en el cuadro de cotización
  var cant_filas_cuadro = fila_fin - inicio_ctz + 1;
  var contenido_ctz = obj_hoja.getRange(inicio_ctz, 1, cant_filas_cuadro, 6).getValues();
  
  // Se ubica el texto que precede al cuadro de cotización, en el documento
  var range = body.findText("A continuación presento la cotización solicitada.");

  
  var ele = range.getElement();
    if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
      // Se ubica el lugar para el cuadro de cotización
      var offset = body.getChildIndex(ele.getParent());
      // Se inserta el cuadro de cotizacion
      var tabla_ctz = body.insertTable(offset + 2, contenido_ctz);
      // Se especifica el estilo común para todo el cuadro
      var estilo_tabla = {};
      estilo_tabla[DocumentApp.Attribute.BORDER_WIDTH] = 0; 
      estilo_tabla[DocumentApp.Attribute.FONT_SIZE] = 10;
      estilo_tabla[DocumentApp.Attribute.SPACING_AFTER] = 0;
      estilo_tabla[DocumentApp.Attribute.LINE_SPACING] = 0;
      
      // Se asigna el estilo común para todo el cuadro
      tabla_ctz.setAttributes(estilo_tabla);
      tabla_ctz.setColumnWidth(0, 35);
      tabla_ctz.setColumnWidth(1, 60);
      tabla_ctz.setColumnWidth(2, 205);
      tabla_ctz.setColumnWidth(4, 60);
      
      // Se especifica el estilo para los encabezados
      var estilo_encabezado = {};
      estilo_encabezado[DocumentApp.Attribute.BOLD] = true;

      // Se asigna el estilo para los encabezados
      var encabezado = tabla_ctz.getRow(0);
      encabezado.setAttributes(estilo_encabezado);
    
      // Se cuenta la cantidad de filas del cuadro de cotización
      var cant_filas = fila_fin - inicio_ctz;
    
      // Se especifica el estilo para los totales
      var estilo_totales = {};
      estilo_totales[DocumentApp.Attribute.BOLD] = true;
      
      // Se asigna el estilo para los totales
      tabla_ctz.getRow(cant_filas-2).setAttributes(estilo_totales);
      tabla_ctz.getRow(cant_filas-1).setAttributes(estilo_totales);
      tabla_ctz.getRow(cant_filas).setAttributes(estilo_totales);
      
      var estilo_derecha = {};
      estilo_derecha[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
    
      var estilo_centrado = {};
      estilo_centrado[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      
      for (var i = 0; i <= cant_filas; i++){
        if(i<cant_filas-2){
          tabla_ctz.getRow(i).getCell(3).getChild(0).asParagraph().setAttributes(estilo_centrado);
          tabla_ctz.getRow(i).getCell(4).getChild(0).asParagraph().setAttributes(estilo_centrado);
          tabla_ctz.getRow(i).getCell(5).getChild(0).asParagraph().setAttributes(estilo_derecha);
        }else{
          tabla_ctz.getRow(i).getCell(4).getChild(0).asParagraph().setAttributes(estilo_derecha);
          tabla_ctz.getRow(i).getCell(5).getChild(0).asParagraph().setAttributes(estilo_derecha);
        }
      } 
    

// Falta colocar el borde por debajo a encabezado y final de página
      // ENSAYOS
      
      /* var atts = tabla_ctz.getAttributes();
      for (var att in atts) {
        Logger.log(att + ":" + atts[att]);
      } */
    
    anunciar('Ya se creó el documento de cotización. Puede buscarlo en gdrive con el nombre de "'+ nombre_doc + '"');
  }else{
    anunciar('En la hoja "Cotización" puede crear el documento de cotización');
  }
  
  
}

