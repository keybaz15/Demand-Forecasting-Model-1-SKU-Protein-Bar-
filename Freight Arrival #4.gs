function populateColumnD() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var inventoryOrdersSheet = spreadsheet.getSheetByName("Inventory Orders");
  var variablesSheet = spreadsheet.getSheetByName("Variables");

  // Get the range of data in columns B and E, get variables tab cells
  var lastRow = inventoryOrdersSheet.getLastRow();
  var dataRange = inventoryOrdersSheet.getRange(2, 2, lastRow - 1, 4); // Columns B and E
  var data = dataRange.getValues();
  var cellFreightMaxDailyDeparture = variablesSheet.getRange("FreightMaxDailyDeparture").getValue(); // Get value from Variables sheet
  var cellFreightToDC1Days = variablesSheet.getRange("FreightToDC1Days").getValue(); // Get value from Variables sheet
  var cellFreightToDC2Days = variablesSheet.getRange("FreightToDC2Days").getValue(); // Get value from Variables sheet
  var cellWarehouse1 = variablesSheet.getRange("Warehouse1").getValue(); // Get value from Variables sheet
  var cellWarehouse2 = variablesSheet.getRange("Warehouse2").getValue(); // Get value from Variables sheet
  var eligibleWarehouseRecievingDays = variablesSheet.getRange("EligibleWarehouseRecievingDays").getValue(); // Get value from Variables sheet

  var shippingDateData = [];

  for (var i = 0; i < data.length; i++) {
    var warehouse = data[i][0]; // Column B
    var freightArrivalDate = data[i][3]; // Column E

    if (freightArrivalDate && (warehouse === cellWarehouse1 || warehouse === cellWarehouse2)) {
      // Add days based on the warehouse (cellWarehouse1 or cellWarehouse2)
      var addDays = warehouse === cellWarehouse1 ? cellFreightToDC1Days : cellFreightToDC2Days;
      var shippingDate = new Date(freightArrivalDate);
      shippingDate.setDate(shippingDate.getDate() + addDays);

      // Find the nearest eligible receiving day
      shippingDate = findNearestEligibleDay(shippingDate, eligibleWarehouseRecievingDays);

      shippingDateData.push([shippingDate]);
    } else {
      shippingDateData.push(['']);
    }
  }

  // Set the values in the "Shipping Date" column (D)
  if (shippingDateData.length > 0) {
    inventoryOrdersSheet.getRange(2, 4, shippingDateData.length, 1).setValues(shippingDateData);
  }
}

function findNearestEligibleDay(date, eligibleDays) {
  // Split the eligible days into an array
  var daysArray = eligibleDays.split(',').map(function(day) {
    return day.trim(); // Remove extra whitespace
  });

  // Find the day index of the input date (0 = Sunday, 1 = Monday, ..., 6 = Saturday)
  var inputDayIndex = date.getDay();

  // Iterate to find the nearest eligible day
  for (var i = 0; i < 7; i++) {
    // Calculate the index of the next day
    var nextDayIndex = (inputDayIndex + i) % 7;

    // Get the name of the next day (e.g., "Monday")
    var nextDayName = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][nextDayIndex];

    // Check if the nextDayName is in the list of eligible days
    if (daysArray.includes(nextDayName)) {
      // Calculate the number of days to add to reach the eligible day
      var daysToAdd = i;

      // Adjust the date accordingly
      date.setDate(date.getDate() + daysToAdd);
      return date;
    }
  }

  // If no eligible day was found, return the input date
  return date;
}
