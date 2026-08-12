/**
 * Expone cualquier pestaña de esta Google Sheet como JSON.
 * Uso: la URL de despliegue + "?sheet=precios" (o mix, omie, analisis)
 */
function doGet(e) {
  var nombreHoja = (e && e.parameter && e.parameter.sheet) || "precios";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombreHoja);

  if (!hoja) {
    var error = { error: "Hoja no encontrada: " + nombreHoja };
    return ContentService.createTextOutput(JSON.stringify(error))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var cabecera = datos[0];
  var filas = datos.slice(1).map(function (fila) {
    var objeto = {};
    cabecera.forEach(function (nombreColumna, i) {
      var valor = fila[i];
      if (valor instanceof Date) {
        valor = Utilities.formatDate(valor, "Europe/Madrid", "yyyy-MM-dd HH:mm:ss");
      }
      objeto[nombreColumna] = valor;
    });
    return objeto;
  });

  return ContentService.createTextOutput(JSON.stringify(filas))
    .setMimeType(ContentService.MimeType.JSON);
}
