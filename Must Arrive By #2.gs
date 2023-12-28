function getSheetByGID(gid) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == gid) {
      return sheets[i];
    }
  }
  return null; // Return null if no sheet with the given GID is found
}

function populateMustArriveByColumn() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Replace with the actual GIDs of your sheets
  var monthlyInventoryGID = 0; 
  var inventoryOrdersGID = 466959541;

  var sheet = getSheetByGID(monthlyInventoryGID);
  var inventoryOrdersSheet = getSheetByGID(inventoryOrdersGID);

  // Get columns for DC1ContainersNeeded and DC2ContainersNeeded
  var dc1Column = spreadsheet.getRangeByName("DC1ContainersNeeded").getColumn();
  var dc2Column = spreadsheet.getRangeByName("DC2ContainersNeeded").getColumn();

  var dataDC1 = sheet.getRange(25, dc1Column, 76).getValues(); // Rows 25 to 100 in DC1ContainersNeeded column
  var dataDC2 = sheet.getRange(25, dc2Column, 76).getValues(); // Rows 25 to 100 in DC2ContainersNeeded column

  var cellReferences = [];
  for (var i = 0; i < Math.max(dataDC1.length, dataDC2.length); i++) {
    if (i < dataDC1.length && !isNaN(dataDC1[i][0])) {
      cellReferences.push({ col: dc1Column, row: i + 25 });
    }
    if (i < dataDC2.length && !isNaN(dataDC2[i][0])) {
      cellReferences.push({ col: dc2Column, row: i + 25 });
    }
  }

  var datesToPopulate = [];
  var headers = inventoryOrdersSheet.getRange(1, 1, 1, inventoryOrdersSheet.getLastColumn()).getValues()[0];
  var mustArriveByColIndex = headers.indexOf("Must Arrive By") + 1;

  for (var i = 0; i < cellReferences.length; i++) {
    var cellRef = cellReferences[i];
    var cellValue = sheet.getRange(cellRef.row, cellRef.col).getValue();
    if (cellValue > 0) {
      var dateValue = sheet.getRange(cellRef.row, 2).getValue(); // Assuming date values are in column B
      for (var j = 0; j < cellValue; j++) {
        var adjustedDate = adjustToPreviousBusinessDay(dateValue);
        datesToPopulate.push([adjustedDate]);
      }
    }
  }

  if (mustArriveByColIndex > 0) {
    inventoryOrdersSheet.getRange(2, mustArriveByColIndex, inventoryOrdersSheet.getLastRow() - 1, 1).clear();
    inventoryOrdersSheet.getRange(2, mustArriveByColIndex, datesToPopulate.length, 1).setValues(datesToPopulate);
  } else {
    Logger.log("Column 'Must Arrive By' not found.");
  }
}

function adjustToPreviousBusinessDay(date) {
  var dayOfWeek = date.getDay();
  if (dayOfWeek === 0) { // Sunday
    date.setDate(date.getDate() - 2); // Move to the previous Friday
  } else if (dayOfWeek === 6) { // Saturday
    date.setDate(date.getDate() - 1); // Move to the previous Friday
  }
  return date;
}
