function populateFreightArrival() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var inventoryOrdersSheet = spreadsheet.getSheetByName("Inventory Orders");
  var variablesSheet = spreadsheet.getSheetByName("Variables");
  
  // Get values from cells in 'Variables' 
  var cellEligibleFreightDepartureDays = variablesSheet.getRange("EligibleFreightDepartureDays").getValue();
  var cellFreightMaxDailyDeparture = variablesSheet.getRange("FreightMaxDailyDeparture").getValue();
  var cellFreightToDC1Days = variablesSheet.getRange("FreightToDC1Days").getValue();
  var cellFreightToDC2Days = variablesSheet.getRange("FreightToDC2Days").getValue();
  var cellWarehouse1 = variablesSheet.getRange("Warehouse1").getValue();
  var cellWarehouse2 = variablesSheet.getRange("Warehouse2").getValue();

  // Get the value from cell 'Variables' tab, EligibleFreightDepartureDays, containing days of the week
  var cellE18 = variablesSheet.getRange("EligibleFreightDepartureDays").getValue();
  
  // Split the cellE18 value into an array of days
  var daysOfWeek = cellE18.split(',').map(function(day) {
    return day.trim(); // Remove leading/trailing spaces
  });

  // Get the range of data in columns B and C
  var lastRow = inventoryOrdersSheet.getLastRow();
  var dataRange = inventoryOrdersSheet.getRange(2, 2, lastRow - 1, 2); // Columns B and C
  var data = dataRange.getValues();

  var freightArrivalData = [];
  var dateCount = {}; // Object to keep track of the count of each date

  for (var i = 0; i < data.length; i++) {
    var warehouse = data[i][0]; // Column B
    var mustArriveByDate = data[i][1]; // Column C

    if (mustArriveByDate && (warehouse === cellWarehouse1 || warehouse === cellWarehouse2)) {
      var subtractDays = warehouse === cellWarehouse1 ? cellFreightToDC1Days : cellFreightToDC2Days;
      var targetDate = new Date(mustArriveByDate);
      var arrivalDate = new Date(targetDate.setDate(targetDate.getDate() - subtractDays));

      // Find a suitable date that is at least FreightToDC1Days or FreightToDC2Days earlier and not used more than FreightMaxDailyDeparture times
      arrivalDate = findSuitableDate(arrivalDate, subtractDays, dateCount, cellFreightMaxDailyDeparture, daysOfWeek);

      freightArrivalData.push([arrivalDate]);
    } else {
      freightArrivalData.push(['']);
    }
  }

  // Set the values in the "Freight Arrival" column (E)
  if (freightArrivalData.length > 0) {
    inventoryOrdersSheet.getRange(2, 5, freightArrivalData.length, 1).setValues(freightArrivalData);
  }
}

function findSuitableDate(targetDate, subtractDays, dateCount, maxCount, daysOfWeek) {
  while (true) {
    var formattedDate = formatDate(targetDate);
    if (targetDate <= new Date(formattedDate) && (!dateCount[formattedDate] || dateCount[formattedDate] < maxCount) && isDayOfWeek(targetDate, daysOfWeek)) {
      // Increment the count for this date
      dateCount[formattedDate] = (dateCount[formattedDate] || 0) + 1;
      return targetDate;
    }
    // Move back a day
    targetDate.setDate(targetDate.getDate() - 1);
  }
}

function isDayOfWeek(date, daysOfWeek) {
  // Check if the day of the week matches any in the array
  var day = date.toLocaleDateString('en-US', { weekday: 'long' });
  return daysOfWeek.includes(day);
}

function formatDate(date) {
  // Format the date as a string for easy comparison and tracking
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
}
