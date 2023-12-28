function populateInventoryOrders() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  function getSheetByGID(gid) {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() == gid) {
        return sheets[i];
      }
    }
    return null; // Return null if no sheet with the given GID is found
  }

  var monthlyInventoryGID = 0; // GID for "Monthly Inventory"
  var inventoryOrdersGID = 466959541; // GID for "Inventory Orders"

  var monthlyInventorySheet = getSheetByGID(monthlyInventoryGID);
  var inventoryOrdersSheet = getSheetByGID(inventoryOrdersGID);

  var warehouse1 = spreadsheet.getRangeByName("Warehouse1").getValue();
  var warehouse2 = spreadsheet.getRangeByName("Warehouse2").getValue();

  if (!monthlyInventorySheet || !inventoryOrdersSheet) {
    Logger.log("Sheet not found. Check the sheet names.");
    return;
  }

  // Get columns for DC1ContainersNeeded and DC2ContainersNeeded
  var dc1Column = spreadsheet.getRangeByName("DC1ContainersNeeded").getColumn();
  var dc2Column = spreadsheet.getRangeByName("DC2ContainersNeeded").getColumn();

  var data = [];
  var distributionCenterHeaderRange = spreadsheet.getRangeByName("DistributionCenterHeader");
  if (!distributionCenterHeaderRange) {
    Logger.log("Named range 'DistributionCenterHeader' not found.");
    return;
  }
  var distributionCenterColIndex = distributionCenterHeaderRange.getColumn();

  for (var row = 25; row <= 100; row++) {
    var dc1Value = Number(monthlyInventorySheet.getRange(row, dc1Column).getValue());
    var dc2Value = Number(monthlyInventorySheet.getRange(row, dc2Column).getValue());

    [dc1Value, dc2Value].forEach(function (value, index) {
      if (isNaN(value)) {
        Logger.log("Non-numeric value in row " + row);
        return;
      }

      var warehouse = (index === 0) ? warehouse1 : warehouse2;
      for (var j = 0; j < value; j++) {
        data.push([warehouse]);
      }
    });
  }

  if (data.length > 0) {
    // Clear existing content in the Distribution Center column from row 2 to 1000
    inventoryOrdersSheet.getRange(2, distributionCenterColIndex, 999).clearContent();

    // Populate the Distribution Center column with the new data
    inventoryOrdersSheet.getRange(2, distributionCenterColIndex, data.length, 1).setValues(data);
  } else {
    Logger.log("No data to populate.");
  }
}
