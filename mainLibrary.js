// Library Project Code


// Function to log messages to a specific sheet
function logMessage(message) {
  var logSheetId = '1kXeqPQAPCcplYM7OQP7mA20uEnWK2OzqT02vS6DXiXg'; // Replace with your actual log sheet ID
  var logSheet = SpreadsheetApp.openById(logSheetId).getSheetByName('Logs'); // Replace 'Logs' with the name of your log sheet
  logSheet.appendRow([new Date(), message]);
}

function processPropertyUpdates(e) {
  var sourceSheet = e.source.getActiveSheet();
  var sheetId = e.source.getId();
  
  var partnerSpreadsheetMap = {
    'JJ': '1ql1F9fTs2Tv6kZr53jDG8GKad-K2_5FT2gus-06w06E',
    'Disponible': '1IbrbbGNGzPIPVIT11lUUQarzVKD361rl_3sl_M74JeM',
    'VIVE': '1TNDsR_l9ycfCvdO-ev4whcBZwNT5hqYvSCRlzFVbwG4',
    'TIKO': '1nJODCLYfgSiJWjOIkjgnoJugEgQKg9ZNeaOmsmovATM',
    'SECOES': '1npAX1Z7kz2y3Z7Z90j8W4kNKy0gbPzCaZY4FmrGgcbo',
    'PERCENT': '1doZEA-ST20lwzUUc89sRuiBA1y0oZWrrDi7h_-qUxhY',
    'MARIO GARCÍA': '1PVx7s5uShdMsms-hL61XxcPA3fZRAUSmW9ncUz9KGP0',
    'LANDA': '1ugGiY-B3c91T9PcQ_Sj3TniyYPEZwHLKFJ_M1ra-AP4',
    'INMOTASA': '1C1GKy2BWeoz9L2HdUOuzI6LuD4ej4Cz6HhKbY7X_Uck',
    'FUSION INMO': '1rMOoO4yA_ofB8221rQIxLybF4xGVu_TzovmV8YH2DFc',
    'CASTALIA': '1TXVHdfllO9u3qj2kKI_D1agGd6Ln72focx_R9L6Sk4o',
    'BOUTIQUE': '1UtebJcvQqCUGf0MzbCTqZjZZHYAoTasn_gcp0SiDWR0',
    'ADRIAN BRIEGA': '1s9d78DWL_Ozemo8N-HAKn4rfV0NpauuXUx-e8eiEvdk',
    'OTROS': 'SpreadsheetID14',
    '-': null // Represents no partner selected
  };

  var headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

  var idColumnIndex = headers.indexOf("Property URL") + 1;
  var statusColumnIndexSource = headers.indexOf("PropHero - Status") + 1;
  var propHeroCommentsColumnIndex = headers.indexOf("PropHero - Comments") + 1;
  var partnerListColumnIndex = headers.indexOf("Partners List") + 1;
  var priceColumnIndex = headers.indexOf("Price (k)") + 1;
  var propertyDetailsColumnIndex = headers.indexOf("Property Details") + 1;
  var filterColumnIndex = headers.indexOf("Filter") + 1;
  var estimatedYieldColumnIndex = headers.indexOf("Estimated Yield (internal)") + 1;
  var yieldValidatedColumnIndex = headers.indexOf("Yield Validated") + 1;

  var row = e.range.getRow();
  var column = e.range.getColumn();
  var data = sourceSheet.getRange(row, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  var partnerValue = data[partnerListColumnIndex - 1];
  var propertyId = data[idColumnIndex - 1];
  var statusValue = data[statusColumnIndexSource - 1];

  // Only trigger if the status is "Enviado al partner"
  if (statusValue !== "Enviado al partner") {
    return;
  }

  if ((column == partnerListColumnIndex && partnerValue != "" && partnerValue != "-") || (column == propHeroCommentsColumnIndex && e.value != "")) {
    var targetSpreadsheetID = partnerSpreadsheetMap[partnerValue];
    if (!targetSpreadsheetID) {
      return;
    }

    var targetSheetName = 'Scraper';
    var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetID);
    var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
    if (!targetSheet) {
      return;
    }

    var targetRow = findRowById(targetSheet, propertyId, "Property URL");
    if (targetRow == -1) {
      var targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
      var statusColumnIndexTarget = targetHeaders.indexOf("PropHero - Status") + 1;
      var propHeroCommentsColumnIndexTarget = targetHeaders.indexOf("PropHero - Comments") + 1;

      var newRow = [];
      for (var i = 0; i < targetHeaders.length; i++) {
        var columnName = targetHeaders[i];
        if (columnName == "PropHero - Status") {
          newRow.push("Enviado al partner");
        } else if (columnName == "PropHero - Comments") {
          newRow.push(data[propHeroCommentsColumnIndex - 1]);
        } else if (columnName == "Status Partner") {
          newRow.push("Por revisar");
        } else if (columnName == "Estimated Yield (internal)") {
          newRow.push(data[estimatedYieldColumnIndex - 1]);
        } else if (columnName == "Yield Validated") {
          newRow.push(data[yieldValidatedColumnIndex - 1]);
        } else {
          newRow.push(data[headers.indexOf(columnName)]);
        }
      }

      targetSheet.appendRow(newRow);

      var propertyUrl = data[idColumnIndex - 1];
      var price = data[priceColumnIndex - 1];
      var filter = data[filterColumnIndex - 1];
      var propertyDetails = data[propertyDetailsColumnIndex - 1];

      var recipientEmails = partnerEmailMap[partnerValue];
      if (!recipientEmails) {
        return;
      }

      var emailSubject = '🌞 Nueva Propiedad Añadida 🌞';
      var emailBody = `Detalles:

      - 🗺️ Zona: ${filter}
      - ✍️ Details: ${propertyDetails}
      - 💰 Precio: €${price},000

  ‼️ Importante actualizar el estado en el scraper => ${targetSpreadsheet.getUrl()}

      `;

      try {
        MailApp.sendEmail(recipientEmails, emailSubject, emailBody);
      } catch (error) {
        logMessage('Failed to send email: ' + error.toString());
      }
    }

    var cell = sourceSheet.getRange(row, partnerListColumnIndex);
    cell.clearComment();
    cell.setBackground(null);
  }

  if (column == propHeroCommentsColumnIndex && e.value != "") {
    var targetSpreadsheetID = partnerSpreadsheetMap[partnerValue];
    if (targetSpreadsheetID) {
      var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetID);
      var targetSheet = targetSpreadsheet.getSheetByName('Scraper');
      if (targetSheet) {
        var targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
        var targetRow = findRowById(targetSheet, propertyId, "Property URL");
        var propHeroCommentsColumnIndexTarget = targetHeaders.indexOf("PropHero - Comments") + 1;

        if (targetRow != -1) {
          var propHeroComments = sourceSheet.getRange(row, propHeroCommentsColumnIndex).getValue();
          targetSheet.getRange(targetRow, propHeroCommentsColumnIndexTarget).setValue(propHeroComments);
        }
      }
    }
  }
}

// Helper function to create a new row for the target sheet
function createNewRow(data, headers, targetSheet) {
  logMessage("createNewRow called");
  var targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  var newRow = [];
  for (var i = 0; i < targetHeaders.length; i++) {
    var columnName = targetHeaders[i];
    if (columnName == "PropHero - Status") {
      newRow.push("Enviado al partner");
    } else if (columnName == "PropHero - Comments") {
      newRow.push(data[headers.indexOf("PropHero - Comments")]);
    } else if (columnName == "Status Partner") {
      newRow.push("Por revisar");
    } else if (columnName == "Estimated Yield (internal)") {
      newRow.push(data[headers.indexOf("Estimated Yield (internal)")]);
    } else if (columnName == "Yield Validated") {
      newRow.push(data[headers.indexOf("Yield Validated")]);
    } else {
      newRow.push(data[headers.indexOf(columnName)]);
    }
  }
  logMessage("New row created: " + newRow);
  return newRow;
}

// Function to move a row within Partner Sheet to the "Prospectos" sheet (transposed)
function moveToProspectos(sheetId, rowIndex) {
  logMessage("moveToProspectos called with sheetId: " + sheetId + ", rowIndex: " + rowIndex);
  var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  var row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Move to "Prospectos" sheet
  var prospectosSheet = SpreadsheetApp.openById(sheetId).getSheetByName('Prospectos');
  var prospectosHeaders = prospectosSheet.getRange(1, 1, prospectosSheet.getLastRow(), 1).getValues().flat();
  var transposedRow = transposeRowWithMatching(row, sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0], prospectosHeaders);
  
  // Find the "Origin" column and set its value to "Scraper"
  var originColumnIndex = prospectosHeaders.indexOf("Origin");
  if (originColumnIndex !== -1) {
    transposedRow[originColumnIndex] = ["Scraper"];
  }

  var lastColumn = prospectosSheet.getLastColumn() + 1;
  var range = prospectosSheet.getRange(1, lastColumn, transposedRow.length, 1);
  range.setValues(transposedRow);

  // Optionally, you can remove the row from the original sheet if needed
  sheet.deleteRow(rowIndex);
  logMessage("Row moved to Prospectos and transposed with Origin set to Scraper");
}

// Function to move a column from "Prospectos" to another sheet within the partner sheets
function moveColumnToOtherSheets(sheetId, sheetName, columnIndex, targetSheetName) {
  logMessage("moveColumnToOtherSheets called with sheetId: " + sheetId + ", sheetName: " + sheetName + ", columnIndex: " + columnIndex + ", targetSheetName: " + targetSheetName);
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  
  // Select all values of the column where the edit was made
  var columnData = sheet.getRange(1, columnIndex, sheet.getLastRow(), 1).getValues();
  logMessage("Column data fetched: " + JSON.stringify(columnData));

  // Verify if the column data is not empty
  if (!columnData || columnData.length === 0) {
    logMessage("Error: Column data is empty or undefined.");
    return;
  }

  // Move to target sheet
  var targetSheet = SpreadsheetApp.openById(sheetId).getSheetByName(targetSheetName);
  if (!targetSheet) {
    logMessage("Error: Target sheet " + targetSheetName + " not found.");
    return;
  }

  var lastColumn = targetSheet.getLastColumn() + 1;

  // Insert the new column into the target sheet in one batch
  targetSheet.getRange(1, lastColumn, columnData.length, 1).setValues(columnData);
  logMessage("Inserted column data into target sheet");

  // Optionally, you can clear the column from the original sheet if needed
  sheet.getRange(1, columnIndex, sheet.getLastRow(), 1).clearContent();
  logMessage("Column moved within sheet to " + targetSheetName);
}

// Function to match row values with headers for column-structured sheets
function matchRowWithHeadersForColumnStructure(row, sourceHeaders, targetHeaders) {
  logMessage("matchRowWithHeadersForColumnStructure called");
  var newRow = [];
  for (var i = 0; i < targetHeaders.length; i++) {
    var columnName = targetHeaders[i];
    var sourceIndex = sourceHeaders.indexOf(columnName);
    if (sourceIndex !== -1) {
      newRow.push(row[sourceIndex]);
    } else {
      newRow.push('');
    }
  }
  logMessage("Row matched with headers: " + JSON.stringify(newRow));
  return newRow;
}

// Function to transpose a row with matching column names
function transposeRowWithMatching(row, sourceHeaders, targetHeaders) {
  logMessage("transposeRowWithMatching called");
  var transposed = [];
  for (var i = 0; i < targetHeaders.length; i++) {
    var columnName = targetHeaders[i];
    var sourceIndex = sourceHeaders.indexOf(columnName);
    if (sourceIndex !== -1) {
      transposed.push([row[sourceIndex]]);
    } else {
      transposed.push(['']);
    }
  }
  logMessage("Row transposed with matching: " + transposed);
  return transposed;
}

// Function to receive updates from Partner Sheets and send them back to the Master Sheet
function receiveUpdatesFromPartner(row, masterSheetId) {
  logMessage("receiveUpdatesFromPartner called with masterSheetId: " + masterSheetId);
  var masterSheet = SpreadsheetApp.openById(masterSheetId).getSheetByName('spain');
  var data = masterSheet.getDataRange().getValues();

  // Find the matching row in the Master Sheet to update
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === row[0]) { // Assuming the first column is a unique identifier
      var range = masterSheet.getRange(i + 1, 1, 1, row.length);
      range.setValues([row]);
      logMessage("Row updated in Master Sheet: " + row);
      break;
    }
  }
}

// Helper function to find a row by ID in a sheet
function findRowById(sheet, id, idColumnName) {
  logMessage("findRowById called with id: " + id + ", idColumnName: " + idColumnName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idColumnIndex = headers.indexOf(idColumnName) + 1;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { // Start from 1 to skip header row
    if (data[i][idColumnIndex - 1] == id) {
      logMessage("Row found with id: " + id + " at index: " + (i + 1));
      return i + 1;
    }
  }
  logMessage("Row not found with id: " + id);
  return -1;
}
