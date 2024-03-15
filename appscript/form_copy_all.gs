
// Guardar celdas
function Guardar(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var formulario = spreadsheet.getSheetByName("Formulario");
  var datos = spreadsheet.getSheetByName("Datos");
  var lastRow = datos.getLastRow();

  formulario.getRange("B5:H50").copyValuesToRange(datos, 1, 6, lastRow+1, lastRow+4)
}

function Limpiar(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var formulario = spreadsheet.getSheetByName("Formulario");
  formulario.getRange("B5:H50").clear()
}
