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
    'JJ': '1ql1F9fTs2Tv6kZr53jDG8GKad-K2_5FT2gus-06w06E', //
    'Disponible': '1IbrbbGNGzPIPVIT11lUUQarzVKD361rl_3sl_M74JeM',
    'VIVE': '1TNDsR_l9ycfCvdO-ev4whcBZwNT5hqYvSCRlzFVbwG4',
    'TIKO': '1nJODCLYfgSiJWjOIkjgnoJugEgQKg9ZNeaOmsmovATM',
    'SECOES': '1npAX1Z7kz2y3Z7Z90j8W4kNKy0gbPzCaZY4FmrGgcbo',
    'PERCENT': '1doZEA-ST20lwzUUc89sRuiBA1y0oZWrrDi7h_-qUxhY',
    'MARIO GARCÍA': '1PVx7s5uShdMsms-hL61XxcPA3fZRAUSmW9ncUz9KGP0',
    'LANDA': '1ugGiY-B3c91T9PcQ_Sj3TniyYPEZwHLKFJ_M1ra-AP4',
    'INMOTASA': '1G4MvgXY8okdnc2L3r4U9GF-ymssQxAJVtC2x2_v-ADg', //
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
function moveToProspectos(sheetId, rowIndex, sheetName) {
  logMessage("moveToProspectos called with sheetId: " + sheetId + ", rowIndex: " + rowIndex);
  var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  var row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Move to "Prospectos" sheet
  var prospectosSheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
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
function receiveUpdatesFromPartner(sheetName, row, masterSheetId) {
  const editedRange = row.range;
  const sheet = editedRange.getSheet();
  const sourceSpreadsheetId = masterSheetId; // Master Sheet ID passed as argument
  const sourceSheetName = "spain"; // Name of the source sheet

  const targetSpreadsheet = SpreadsheetApp.openById(row.spreadsheetId);
  const targetSheet = targetSpreadsheet.getSheetByName(sheetName);
  const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

  const targetHeaders = sheetName === "Scraper" ? 
    targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0] : 
    targetSheet.getRange(1, 1, targetSheet.getLastRow(), 1).getValues().flat();
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

  const idColumnIndexTarget = targetHeaders.indexOf("Property URL");
  const statusPartnerColumnIndexTarget = targetHeaders.indexOf("Status Partner");
  const statusPartnerIIColumnIndexTarget = targetHeaders.indexOf("Status II (post-visita)");
  const commentsPartnerColumnIndexTarget = targetHeaders.indexOf("Comments Partner");

  const idColumnIndexSource = sourceHeaders.indexOf("Property URL");
  const statusPartnerColumnIndexSource = sourceHeaders.indexOf("Status Partner");
  const statusPartnerIIColumnIndexSource = sourceHeaders.indexOf("Status II (post-visita)");
  const commentsPartnerColumnIndexSource = sourceHeaders.indexOf("Comments Partner");

  const rowNum = editedRange.getRow();
  const colNum = editedRange.getColumn();
  const newValue = editedRange.getValue(); // Get the new value to sync

  logMessage(`Sheet name: ${sheetName}`);
  logMessage(`Row: ${rowNum}, Column: ${colNum}, New value: ${newValue}`);

  if (sheetName === "Scraper") {
    // Row-based structure
    const id = targetSheet.getRange(rowNum, idColumnIndexTarget + 1).getValue(); // Get the unique ID of the edited row
    logMessage(`ID from Scraper sheet: ${id}`);
    if (!id) return;

    // Find the row with the same ID in the source sheet
    const sourceDataRange = sourceSheet.getDataRange();
    const sourceValues = sourceDataRange.getValues();
    for (let i = 1; i < sourceValues.length; i++) { // Start from 1 to skip header row
      if (sourceValues[i][idColumnIndexSource] === id) {
        // Update the corresponding source sheet column
        if (colNum - 1 === statusPartnerColumnIndexTarget) {
          logMessage(`Updating Status Partner in row ${i + 1}`);
          sourceSheet.getRange(i + 1, statusPartnerColumnIndexSource + 1).setValue(newValue);
        } else if (colNum - 1 === statusPartnerIIColumnIndexTarget) {
          logMessage(`Updating Status II (post-visita) in row ${i + 1}`);
          sourceSheet.getRange(i + 1, statusPartnerIIColumnIndexSource + 1).setValue(newValue);
        } else if (colNum - 1 === commentsPartnerColumnIndexTarget) {
          logMessage(`Updating Comments Partner in row ${i + 1}`);
          sourceSheet.getRange(i + 1, commentsPartnerColumnIndexSource + 1).setValue(newValue);
        }
        break;
      }
    }
  } else {
    // Column-based structure
    const id = targetSheet.getRange(idColumnIndexTarget + 1, colNum).getValue(); // Get the unique ID of the edited column
    logMessage(`ID from ${sheetName} sheet: ${id}`);
    if (!id) {
      logMessage(`No ID found in ${sheetName} at column ${colNum}`);
      return;
    }

    // Find the row with the same ID in the source sheet
    const sourceDataRange = sourceSheet.getDataRange();
    const sourceValues = sourceDataRange.getValues();
    for (let i = 1; i < sourceValues.length; i++) { // Start from 1 to skip header row
      if (sourceValues[i][idColumnIndexSource] === id) {
        logMessage(`Found matching ID at row ${i + 1} in source sheet`);

        try {
          if (rowNum === statusPartnerIIColumnIndexTarget + 1) {
            logMessage(`Updating Status II (post-visita) in row ${i + 1}`);
            sourceSheet.getRange(i + 1, statusPartnerIIColumnIndexSource + 1).setValue(newValue); // Update Status II (post-visita)
          } else if (rowNum === statusPartnerColumnIndexTarget + 1) {
            logMessage(`Updating Status Partner in row ${i + 1}`);
            sourceSheet.getRange(i + 1, statusPartnerColumnIndexSource + 1).setValue(newValue); // Update Status Partner
          } else if (rowNum === commentsPartnerColumnIndexTarget + 1) {
            logMessage(`Updating Comments Partner in row ${i + 1}`);
            sourceSheet.getRange(i + 1, commentsPartnerColumnIndexSource + 1).setValue(newValue); // Update Comments Partner
          }
          logMessage(`Updated value to ${newValue}`);
        } catch (error) {
          logMessage(`Failed to update: ${error.message}`);
        }
        break;
      }
    }
  }
}

// Helper function to find a row by ID in a sheet
function findRowById(sheet, id, idColumnName) {
  logMessage("findRowById called with id: " + id + ", idColumnName: " + idColumnName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  logMessage("Headers in findRowById: " + JSON.stringify(headers));
  var idColumnIndex = headers.indexOf(idColumnName) + 1;
  logMessage("ID Column Index: " + idColumnIndex);
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

function generateId(){
  id = 'ID-' + Math.random().toString(36).substring(2, 11);
  Logger.log("id ...      " + id)
  return id 
}
