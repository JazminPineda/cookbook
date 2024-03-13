
n GetTimeZone()
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
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getRange("B3:F7").getValues();
  for (var i = 0; i < values.length; i++) {
     Logger.log(values[i])
  }

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
