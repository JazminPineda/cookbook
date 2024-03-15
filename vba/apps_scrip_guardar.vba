# Guardar y Limpiar Formulario 
function GetTimeZone()
{
  // Logger.log(Intl.DateTimeFormat().resolvedOptions().timeZone);
  return "America/Buenos_Aires";
}

function GetDate() {
  return Utilities.formatDate(new Date(),GetTimeZone(),"dd/MM/yyyy");
}

function GetValues(range){
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  var hojaFormulario = hojaActiva.getSheetByName("Formulario");
  return hojaFormulario.getRange(range).getValues();
}

function SaveValues(lastRow,length, column, valores){
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  var datos = hojaActiva.getSheetByName("Datos");
  datos.getRange(lastRow+1,column,length,1).setValues(valores);
}

function GardarSociedades(rangoSociedad){
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
  var datos = hojaActiva.getSheetByName("Datos"); // Nombre de hoja donde se almacenan datos
  var lastRow = datos.getLastRow();

  var calidadValores = GetValues(rangoSociedad[1]);
  var cantidad = calidadValores.length;

  datos.getRange(lastRow+1,1,cantidad,1).setValue(GetDate()); // lastrow es la ultima fila + 1 / el numero  columna 1

  SaveValues(lastRow,cantidad,2,calidadValores); // lastrow es la ultima fila + 1 / el numero representa la columna  2

  var calidadValores = GetValues(rangoSociedad[0]);
  SaveValues(lastRow,calidadValores.length,3,calidadValores); // lastrow es la ultima fila + 1 / el numero representa la columna 3

  var calidadValores = GetValues(rangoSociedad[2]);
  SaveValues(lastRow,calidadValores.length,4,calidadValores); // lastrow es la ultima fila + 1 / el numero representa la columna 3

  var calidadValores = GetValues(rangoSociedad[3]);
  SaveValues(lastRow,calidadValores.length,5,calidadValores); // lastrow es la ultima fila + 1 / el numero representa la columna  4

  // var calidadValores = GetValues(rangoSociedad[4]);
  // SaveValues(lastRow,calidadValores.length,6,calidadValores); // lastrow es la ultima fila + 1 / el numero representa la columna 3

  //Limpiar
}

// Guardar celdas
function Guardar(){
  var rangos = {
    bariloche: ["A3:A10", "B3:B10", "G3:G10", "H3:H10"],
    cometa: ["A12:A14", "B12:B14", "G12:G14", "H12:H14"],
    ersa: ["A16:A18","B16:B18", "G16:G18", "H16:H18"],
    pulqui: ["A20:A26","B20:B26", "G20:G26", "H20:H26"],
    jbl: ["A28:A29","B28:B29", "G28:G29", "H28:H29"],
    derudder: ["A33:A72","B33:B72", "G33:G72", "H33:H72"],
    expreso: ["A75:A76","B75:B76", "G75:G76", "H75:H76"],
    junio20: ["A78","B78", "G78", "H78"],
  }
  GardarSociedades(rangos.bariloche);
  GardarSociedades(rangos.cometa);
  GardarSociedades(rangos.ersa);
  GardarSociedades(rangos.pulqui);
  GardarSociedades(rangos.jbl);
  GardarSociedades(rangos.derudder);
  GardarSociedades(rangos.expreso);
  GardarSociedades(rangos.junio20);
}
function Limpiar(){
  var rangos = {
    bariloche: ["G3:G10", "H3:H10", "I3:I10", "J3:J10"],
    cometa: ["G12:G14", "H12:H14", "I12:I14", "J12:J14"],
    ersa: [ "G16:G18", "H16:H18", "I12:I14", "J12:J14"],
    pulqui: ["G20:G26", "H20:H26","I20:I26", "J20:J26"],
    jbl: ["G28:G29", "H28:H29","I28:I29", "J28:J29"],
    derudder: ["G33:G71", "H33:H71","I33:I71", "J33:J71"],
    expreso: ["G73:G74", "H33:H74","I33:I74", "J33:J74"],
    junio20: ["G76", "H76","I76", "J76"],
  }
  LimpiarRango(rangos.bariloche)
  LimpiarRango(rangos.cometa)
  LimpiarRango(rangos.ersa)
  LimpiarRango(rangos.pulqui)
  LimpiarRango(rangos.jbl)
  LimpiarRango(rangos.derudder)
  LimpiarRango(rangos.expreso)
  LimpiarRango(rangos.junio20)
}
function LimpiarRango(rangos){
 var spreadsheet = SpreadsheetApp.getActive();
 var hojaFormulario = spreadsheet.getSheetByName("Formulario");

 for (i=0; i<rangos.length; i++){
    hojaFormulario.getRange(rangos[i]).clearContent();
 }
}
