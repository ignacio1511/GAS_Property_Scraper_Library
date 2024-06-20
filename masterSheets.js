function sendToPartner(e) {
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
    var statusColumnIndex = headers.indexOf("PropHero - Status") + 1;
    var propHeroCommentsColumnIndex = headers.indexOf("PropHero - Comments") + 1;
    var partnerListColumnIndex = headers.indexOf("Partners List") + 1;
  
    PropertyScraperLibrary.logMessage("Edit detected in column: " + range.getColumn() + ", value: " + range.getValue());
  
    var statusValue = sheet.getRange(range.getRow(), statusColumnIndex).getValue();
    
    if (statusValue === "Enviado al partner" && 
        (range.getColumn() === statusColumnIndex || 
         range.getColumn() === propHeroCommentsColumnIndex || 
         range.getColumn() === partnerListColumnIndex)) {
      PropertyScraperLibrary.processPropertyUpdates(e);
    }
  }
  