var app_ctz = SpreadsheetApp.getActiveSpreadsheet();

var id_bd_jfe = "1zqg05_yR9i3lsPWg2BmVLUaKSr99HlZ2n8fGEHSo9Zg";

var plantilla_doc_id = '1cquZ7t0FGgs8f8bkGMhKt4XKXiEe-PeAfGpJ1Aibbyw';

//var bd_jfe = SpreadsheetApp.openById(id_bd_jfe);

//var correo_usuario = Session.getActiveUser().getEmail();*/

// var bd_jfe = FALTA CONETAR CON BD


var nombre_hoja_cotizacion = '01-Cotización';
var nombre_hoja_detalle = '02-Detalle';
var nombre_hoja_plantilla = '03-Plantilla';
var nombre_hoja_partidas = '04-Partidas';
var nombre_hoja_mano_obra = '05-Mano de obra';
var nombre_hoja_materiales = '06-Materiales';
var nombre_hoja_equipos = '07-Equipos';
var nombre_hoja_suministros = '08-Suministros';
var nombre_hoja_bd = 'BD Entera';


var nombre_hoja_bd_jfe_partidas = 'Partidas';
var nombre_hoja_bd_jfe_mano_obra = 'Mano de obra';
var nombre_hoja_bd_jfe_materiales = 'Materiales';
var nombre_hoja_bd_jfe_equipos = 'Equipos';
var nombre_hoja_bd_jfe_suministros = 'Suministros';


var nombre_rango_partidas = 'partidas';
var nombre_rango_mano_obra = 'mano_obra';
var nombre_rango_materiales = 'materiales';
var nombre_rango_equipos = 'equipos';
var nombre_rango_suministros = 'suministros';
var nombre_rango_bd_entera = 'bd_entera';


var fondo_azul = '#CFE2F3';
var fondo_amarillo = '#FAFAD2';

var borde_bajo = {
    top: false,
    left: false,
    bottom: true,
    right: false,
    vertical: false,
    horizontal: false,
    color: '#000000',
    style: SpreadsheetApp.BorderStyle.SOLID
  };

var borde_arriba = {
    top: true,
    left: false,
    bottom: false,
    right: false,
    vertical: false,
    horizontal: false,
    color: '#000000',
    style: SpreadsheetApp.BorderStyle.SOLID
  };

var borde_bajo_medio = {
    top: false,
    left: false,
    bottom: true,
    right: false,
    vertical: false,
    horizontal: false,
    color: '#000000',
    style: SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  };

var borde_arriba_medio = {
    top: true,
    left: false,
    bottom: false,
    right: false,
    vertical: false,
    horizontal: false,
    color: '#000000',
    style: SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  };
