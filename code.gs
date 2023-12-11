function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Cleanup Menu')
    .addItem('Cleanup Data', 'cleanupData')
    .addToUi();
}


function cleanupData() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Replace "aggregate" with the name of your original sheet
  var originalSheet = spreadsheet.getSheetByName("aggregate");
  
  // Get data from the original sheet
  var data = originalSheet.getDataRange().getValues();

  // Batch operation array
  var cleanedData = [];

  // Loop through the data and apply cleanup actions
  for (var i = 2; i < data.length; i++) {
    // Cleanup actions
    var cleanedImages = data[i][13].replace(/[\/+\- ]/g, "");
    var cleanedSubbase = data[i][6].replace(/\s/g, "");
    var cleanedBase = data[i][5].replace(/\s/g, "");
    
var cleanedWarranty = data[i][11].replace(/\s/g, "");

    // Convert price to integer and remove decimal and trailing zeros
    var cleanedPrice = parseInt(data[i][9]);

    var cleanedProfile = isNaN(data[i][3]) ? 0 : Number(data[i][3]);
    var cleanedWidth = Number(data[i][2]);
    var cleanedRim = Number(data[i][4]);

    // Push cleaned data to batch operation array
    cleanedData.push([
      data[i][0], data[i][1], cleanedWidth, cleanedProfile, cleanedRim,
      cleanedBase, cleanedSubbase, data[i][7], data[i][8], cleanedPrice,
      data[i][10], cleanedWarranty, data[i][12], cleanedImages, data[i][14],
      data[i][15]
    ]);
  }

  // Set the timezone to Dubai
  var dubaiTimezone = "Asia/Dubai";

  // Generate timestamp for the sheet name in "dd_mmm_yy_hh_mm" format
  var timestamp = Utilities.formatDate(new Date(), dubaiTimezone, "dd_MMM_yy_HH_mm");

  // Create a new sheet with the timestamp in the name
  var cleanedSheet = spreadsheet.insertSheet("Ali_" + timestamp);

  // Append headers to row 0 in the cleaned sheet
  cleanedSheet.appendRow(originalSheet.getRange(1, 1, 1, originalSheet.getLastColumn()-1).getValues()[0]);

  // Write cleaned data starting from row 1
  cleanedSheet.getRange(2, 1, cleanedData.length, cleanedData[0].length).setValues(cleanedData);

  // Freeze the first row
cleanedSheet.setFrozenRows(1);


  // Set header cell formatting
  var headerRange = cleanedSheet.getRange(1, 1, 1, cleanedSheet.getLastColumn());
  headerRange.setFontWeight("bold")
              .setFontColor("white")
              .setBackgroundColor("#185a9d");

  Logger.log("Cleanup complete!");
}
