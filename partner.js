function onEdit(e) {
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
    PropertyScraperLibrary.logMessage("Edit detected in column: " + range.getColumn() + ", value: " + range.getValue());
  
    if (sheetName === 'Scraper') {
      var statusPartnerColumnIndex = headers.indexOf("Status Partner") + 1;
      if (range.getColumn() === statusPartnerColumnIndex) {
        var status = range.getValue();
        if (status === "Aprobado" || status === "Solicitamos visita") {
          PropertyScraperLibrary.logMessage("Triggering moveToProspectos for row: " + range.getRow());
          PropertyScraperLibrary.moveToProspectos(sheetId, range.getRow());
        } else {
          var row = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
          PropertyScraperLibrary.receiveUpdatesFromPartner(row, 'master-sheet-id'); // Replace 'master-sheet-id' with the actual ID of your Master Sheet
        }
      }
    } else {
      var headersColumn = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
      var statusIIColumnIndex = headersColumn.indexOf("Status II (post-visita)") + 1;
  
      if (range.getRow() === statusIIColumnIndex) {
        var statusII = range.getValue();
        PropertyScraperLibrary.logMessage("Status II edit detected, status: " + statusII);
        if (sheetName === 'Prospectos') {
          if (statusII === "Ofertado" || statusII === "Reservado") {
            PropertyScraperLibrary.logMessage("Triggering moveColumnToOtherSheets to OFERTADAS/RESERVADAS for column: " + range.getColumn());
            PropertyScraperLibrary.moveColumnToOtherSheets(sheetId, sheetName, range.getColumn(), "Ofertadas/Reservadas");
          } else if (statusII === "Descartado") {
            PropertyScraperLibrary.logMessage("Triggering moveColumnToOtherSheets to DESCARTADAS for column: " + range.getColumn());
            PropertyScraperLibrary.moveColumnToOtherSheets(sheetId, sheetName, range.getColumn(), "Descartadas");
          }
        } else if (sheetName === 'Ofertadas/Reservadas') {
          if (statusII === "Cerrado") {
            PropertyScraperLibrary.logMessage("Triggering moveColumnToOtherSheets to CERRADAS for column: " + range.getColumn());
            PropertyScraperLibrary.moveColumnToOtherSheets(sheetId, sheetName, range.getColumn(), "Cerradas");
          } else if (statusII === "Descartado") {
            PropertyScraperLibrary.logMessage("Triggering moveColumnToOtherSheets to DESCARTADAS for column: " + range.getColumn());
            PropertyScraperLibrary.moveColumnToOtherSheets(sheetId, sheetName, range.getColumn(), "Descartadas");
          }
        }
      }
    }
  }
  