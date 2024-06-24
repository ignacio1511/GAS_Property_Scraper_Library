function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  PropertyScraperLibrary.logMessage("Edit detected in column: " + range.getColumn() + ", value: " + range.getValue());

  var statusPartnerColumnIndex = headers.indexOf("Status Partner") + 1;
  var statusIIColumnIndex = headers.indexOf("Status II (post-visita)") + 1;
  var commentsPartnerColumnIndex = headers.indexOf("Comments Partner") + 1;

  PropertyScraperLibrary.logMessage("Outside the IF condition ...");

  if (sheetName === "Scraper") {
    // Row-based structure
    if (range.getColumn() === statusPartnerColumnIndex || range.getColumn() === statusIIColumnIndex || range.getColumn() === commentsPartnerColumnIndex) {
      PropertyScraperLibrary.logMessage("Entering the IF condition for Scraper ...");
      PropertyScraperLibrary.receiveUpdatesFromPartner(sheetName, {
        range: range,
        spreadsheetId: sheetId
      }, '1DE38jM0Ejb-POfpbDpzkkD7X-bgMNArBkMsryt0UmyU');
      PropertyScraperLibrary.logMessage("Function receiveUpdatesFromPartner executed for Scraper!");
    }
  } else {
    // Column-based structure
    var headersColumn = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
    var statusPartnerRowIndex = headersColumn.indexOf("Status Partner") + 1;
    var statusIIRowIndex = headersColumn.indexOf("Status II (post-visita)") + 1;
    var commentsPartnerRowIndex = headersColumn.indexOf("Comments Partner") + 1;

    if (range.getRow() === statusPartnerRowIndex || range.getRow() === statusIIRowIndex || range.getRow() === commentsPartnerRowIndex) {
      PropertyScraperLibrary.logMessage("Entering the IF condition for column-based sheet ...");
      PropertyScraperLibrary.receiveUpdatesFromPartner(sheetName, {
        range: range,
        spreadsheetId: sheetId
      }, '1DE38jM0Ejb-POfpbDpzkkD7X-bgMNArBkMsryt0UmyU');
      PropertyScraperLibrary.logMessage("Function receiveUpdatesFromPartner executed for column-based sheet!");
    }
  }

  if (sheetName === 'Scraper') {
    if (range.getColumn() === statusPartnerColumnIndex) {
      var status = range.getValue();
      if (status === "Aprobado" || status === "Solicitamos visita") {
        PropertyScraperLibrary.logMessage("Triggering moveToProspectos for row: " + range.getRow());
        PropertyScraperLibrary.moveToProspectos(sheetId, range.getRow());
      }
    }
  } else {
    if (range.getRow() === statusIIRowIndex) {
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
