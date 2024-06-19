// Library Project Code


// Function to log messages to a specific sheet
function logMessage(message) {
    var logSheetId = '1kXeqPQAPCcplYM7OQP7mA20uEnWK2OzqT02vS6DXiXg'; // Replace with your actual log sheet ID
    var logSheet = SpreadsheetApp.openById(logSheetId).getSheetByName('Logs'); // Replace 'Logs' with the name of your log sheet
    logSheet.appendRow([new Date(), message]);
  }
  
  // Function to send updates from Master Sheet to Partner Sheets
  function processPropertyUpdates(sheetId) {
    Logger.log("processPropertyUpdates called with sheetId: " + sheetId);
    var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
    var data = sheet.getDataRange().getValues();
  
    var partnerSpreadsheetMap = {
      'JJ': '139ZDyp2K3ijbhyOsjgRmUAVV3fNiHoFThwY-UyfpHoY',
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
  
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idColumnIndex = headers.indexOf("Property URL") + 1;
    var statusColumnIndexSource = headers.indexOf("PropHero - Status") + 1;
    var partnerListColumnIndex = headers.indexOf("Partners List") + 1;
  
    var row = e.range.getRow();
    var column = e.range.getColumn();
    var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    var partnerValue = data[partnerListColumnIndex - 1]; // Get the value in the Partners List column
    var propertyId = data[idColumnIndex - 1]; // Get the unique property ID
  
    Logger.log("partnerValue: " + partnerValue);
    Logger.log("propertyId: " + propertyId);
  
    if (column == statusColumnIndexSource && e.value == "Enviado al partner") {
      if (partnerValue == "" || partnerValue == "-") {
        var cell = sheet.getRange(row, partnerListColumnIndex);
        cell.setComment('Selecciona un partner de la lista para enviar la propiedad');
        cell.setBackground('#FFCCCC'); // Light red background color
        Logger.log('Partners List column is empty. User needs to update it.');
        return;
      }
    }
  
    if (column == partnerListColumnIndex && partnerValue != "" && partnerValue != "-") {
      var targetSpreadsheetID = partnerSpreadsheetMap[partnerValue];
      if (!targetSpreadsheetID) {
        Logger.log('Spreadsheet ID for partner ' + partnerValue + ' not found.');
        return;
      }
  
      var targetSheetName = 'Approved'; // The name of the sheet to use in the partner's spreadsheet
      var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetID);
      var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
      if (!targetSheet) {
        Logger.log('Target sheet for partner ' + partnerValue + ' not found.');
        return;
      }
  
      // Check if the row already exists in the target sheet
      var targetRow = findRowById(targetSheet, propertyId, "Property URL");
      if (targetRow == -1) {
        var newRow = createNewRow(data, headers, targetSheet);
        targetSheet.appendRow(newRow);
        Logger.log("Row appended to partner sheet: " + partnerValue);
      }
  
      var cell = sheet.getRange(row, partnerListColumnIndex);
      cell.clearComment();
      cell.setBackground(null); // Reset to default background
      Logger.log('Partners List column updated. Resetting cell formatting.');
    }
  }
  
  // Helper function to create a new row for the target sheet
  function createNewRow(data, headers, targetSheet) {
    Logger.log("createNewRow called");
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
    Logger.log("New row created: " + newRow);
    return newRow;
  }
  
  // Function to move a row within Partner Sheet to the "Prospectos" sheet (transposed)
  function moveToProspectos(sheetId, rowIndex) {
    Logger.log("moveToProspectos called with sheetId: " + sheetId + ", rowIndex: " + rowIndex);
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
    Logger.log("Row moved to Prospectos and transposed with Origin set to Scraper");
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
    Logger.log("transposeRowWithMatching called");
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
    Logger.log("Row transposed with matching: " + transposed);
    return transposed;
  }
  
  // Function to receive updates from Partner Sheets and send them back to the Master Sheet
  function receiveUpdatesFromPartner(row, masterSheetId) {
    Logger.log("receiveUpdatesFromPartner called with masterSheetId: " + masterSheetId);
    var masterSheet = SpreadsheetApp.openById(masterSheetId).getSheetByName('Master');
    var data = masterSheet.getDataRange().getValues();
  
    // Find the matching row in the Master Sheet to update
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === row[0]) { // Assuming the first column is a unique identifier
        var range = masterSheet.getRange(i + 1, 1, 1, row.length);
        range.setValues([row]);
        Logger.log("Row updated in Master Sheet: " + row);
        break;
      }
    }
  }
  
  // Helper function to find a row by ID in a sheet
  function findRowById(sheet, id, idColumnName) {
    Logger.log("findRowById called with id: " + id + ", idColumnName: " + idColumnName);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idColumnIndex = headers.indexOf(idColumnName) + 1;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { // Start from 1 to skip header row
      if (data[i][idColumnIndex - 1] == id) {
        Logger.log("Row found with id: " + id + " at index: " + (i + 1));
        return i + 1;
      }
    }
    Logger.log("Row not found with id: " + id);
    return -1;
  }
  