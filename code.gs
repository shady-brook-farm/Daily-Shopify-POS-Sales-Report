function onEdit(e) {
  if (e.source.getActiveSheet().getName() === "RawData") {
    processCSVData();
  }
}

function processCSVData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName("RawData");
  var formatted = ss.getSheetByName("FormattedData");
  formatted.clear(); // clears all content & formatting

  // Anchor so your Attendance row + headers stay intact
  var ANCHOR_ROW = 1;
  var ANCHOR_COL = 1;

  // Get data
  var data = raw.getDataRange().getValues();
  if (!data || data.length < 2) return;

  // Normalize headers
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var rows = data.slice(1);

  // Column indices
  var colLocation = headers.indexOf("POS location name");
  var colDate     = headers.indexOf("Day");
  var colVariant  = headers.indexOf("Product variant title at time of sale");
  var colProduct  = headers.indexOf("Product title at time of sale");
  var colQty      = headers.indexOf("Net items sold");
  var colSales    = headers.indexOf("Total sales");

  if (colLocation === -1 || colDate === -1 || colQty === -1) {
    throw new Error("Missing required headers in RawData. Need POS location name, Day, and Net items sold.");
  }
  
  if (colSales === -1) {
    Logger.log("Warning: Total sales column not found, will skip sales calculations");
  }

  var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  // Pivot map and date collection
  var pivot = {};  // { location: { product: { date: qty } } }
  var salesPivot = {}; // { location: { product: { date: sales } } }
  var datesArray = [];
  var seenDates = new Set();
  var productTotals = {}; // { product: { qty: total, sales: total } }

  rows.forEach(function(row) {
    var location = String(row[colLocation] || "").trim();
    if (!location) return;

    var dateObj = new Date(row[colDate]);
    if (!(dateObj instanceof Date) || isNaN(dateObj)) return;

    var dateLabel = Utilities.formatDate(dateObj, tz, "M/d");
    
    // Collect unique dates with their original date objects for proper sorting
    if (!seenDates.has(dateLabel)) {
      seenDates.add(dateLabel);
      datesArray.push({original: dateObj, formatted: dateLabel});
    }

    var product = String(row[colVariant] || row[colProduct] || "Unknown").trim();
    var qty = Number(row[colQty]) || 0;
    var sales = colSales !== -1 ? (Number(row[colSales]) || 0) : 0;

    // Quantity pivot
    if (!pivot[location]) pivot[location] = {};
    if (!pivot[location][product]) pivot[location][product] = {};
    pivot[location][product][dateLabel] = (pivot[location][product][dateLabel] || 0) + qty;

    // Sales pivot (only if sales column exists)
    if (colSales !== -1) {
      if (!salesPivot[location]) salesPivot[location] = {};
      if (!salesPivot[location][product]) salesPivot[location][product] = {};
      salesPivot[location][product][dateLabel] = (salesPivot[location][product][dateLabel] || 0) + sales;
    }

    // Product totals
    if (!productTotals[product]) productTotals[product] = { qty: 0, sales: 0 };
    productTotals[product].qty += qty;
    productTotals[product].sales += sales;
  });

  // Sort by original dates, then extract formatted dates
  var dates = datesArray
    .sort(function(a, b) { return a.original - b.original; })
    .map(function(d) { return d.formatted; });
    
  Logger.log("Dates list: " + JSON.stringify(dates));
  
  // Build header row
  var output = [];
  var headerRow = ["Location", "Product"];
  
  // Add each date as a separate column
  for (var i = 0; i < dates.length; i++) {
    headerRow.push(dates[i]);
  }
  
  Logger.log("Header row length: " + headerRow.length);
  Logger.log("Header row: " + JSON.stringify(headerRow));
  output.push(headerRow);

  // Build grouped rows
  Object.keys(pivot).forEach(function(location) {
    var firstProduct = true;
    Object.keys(pivot[location]).forEach(function(product) {
      var rowOut = [];
      rowOut.push(firstProduct ? location : "");  // only show once
      rowOut.push(product);

      // Add each date's quantity as a separate column
      for (var i = 0; i < dates.length; i++) {
        var qty = pivot[location][product][dates[i]] || 0;
        rowOut.push(qty);
      }

      Logger.log("Row length for " + location + " - " + product + ": " + rowOut.length);
      output.push(rowOut);
      firstProduct = false;
    });

    // blank spacer row between locations (optional)
    var spacerRow = [];
    for (var j = 0; j < headerRow.length; j++) {
      spacerRow.push("");
    }
    output.push(spacerRow);
  });

  // Clear old values in body but keep formatting
  var lastRow = formatted.getLastRow();
  var lastCol = formatted.getLastColumn();
//  if (lastRow >= ANCHOR_ROW && lastCol >= ANCHOR_COL) {
//    formatted.getRange(ANCHOR_ROW, ANCHOR_COL, lastRow - ANCHOR_ROW + 1, lastCol).clearContent();
//  }

  // Debug the output structure before writing
  Logger.log("Total output rows: " + output.length);
  Logger.log("Expected columns: " + headerRow.length);
  
  for (var i = 0; i < Math.min(5, output.length); i++) {
    Logger.log("Row " + (i+1) + " length: " + output[i].length + " | Content: " + JSON.stringify(output[i]));
  }

  // Ensure all rows have the same length as header
  for (var i = 1; i < output.length; i++) {
    while (output[i].length < headerRow.length) {
      output[i].push("");
    }
    if (output[i].length > headerRow.length) {
      output[i] = output[i].slice(0, headerRow.length);
    }
  }

  // Write values - completely reset sheet first
  try {
    // Delete the entire sheet and recreate it to ensure no formatting issues
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Delete existing FormattedData sheet if it exists
    try {
      var existingSheet = spreadsheet.getSheetByName("FormattedData");
      if (existingSheet) {
        spreadsheet.deleteSheet(existingSheet);
      }
    } catch (e) {
      // Sheet doesn't exist, that's fine
    }
    
    // Create a brand new sheet
    var newSheet = spreadsheet.insertSheet("FormattedData");
    
    // Write just the header first to test
    Logger.log("Writing header row...");
    newSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    
    // Write first data row to test
    Logger.log("Writing first data row...");
    if (output.length > 1) {
      newSheet.getRange(2, 1, 1, output[1].length).setValues([output[1]]);
    }
    
    // If those work, write the rest
    Logger.log("Writing remaining data rows...");
    if (output.length > 2) {
      var remainingRows = output.slice(2);
      newSheet.getRange(3, 1, remainingRows.length, headerRow.length).setValues(remainingRows);
    }
    
    // Auto-resize all columns to fit content
    Logger.log("Auto-resizing columns...");
    for (var col = 1; col <= headerRow.length; col++) {
      newSheet.autoResizeColumn(col);
    }
    
    // Add summary section side-by-side below the main data
    Logger.log("Adding summary section...");
    var summaryStartRow = output.length + 3; // Leave some space
    var summaryStartCol = 1;
    var productBreakdownStartCol = 5; // Start product breakdown in column E
    
    // Add "SUMMARY" header (left side)
    newSheet.getRange(summaryStartRow, summaryStartCol).setValue("SUMMARY");
    newSheet.getRange(summaryStartRow, summaryStartCol).setFontWeight("bold").setFontSize(14);
    
    // Add total sales summary (if sales data available)
    if (colSales !== -1) {
      var totalSales = 0;
      var totalQty = 0;
      
      Object.keys(productTotals).forEach(function(product) {
        totalSales += productTotals[product].sales;
        totalQty += productTotals[product].qty;
      });
      
      summaryStartRow += 2;
      newSheet.getRange(summaryStartRow, summaryStartCol).setValue("Total Sales:");
      newSheet.getRange(summaryStartRow, summaryStartCol + 1).setValue("$" + totalSales.toFixed(2));
      newSheet.getRange(summaryStartRow, summaryStartCol).setFontWeight("bold");
      
      summaryStartRow += 1;
      newSheet.getRange(summaryStartRow, summaryStartCol).setValue("Total Quantity Sold:");
      newSheet.getRange(summaryStartRow, summaryStartCol + 1).setValue(totalQty);
      newSheet.getRange(summaryStartRow, summaryStartCol).setFontWeight("bold");
    }
    
    // Add product breakdown header (right side)
    var productBreakdownRow = output.length + 3; // Start at same row as summary
    newSheet.getRange(productBreakdownRow, productBreakdownStartCol).setValue("PRODUCT BREAKDOWN");
    newSheet.getRange(productBreakdownRow, productBreakdownStartCol).setFontWeight("bold").setFontSize(12);
    
    productBreakdownRow += 2;
    var summaryHeaders = ["Product", "Total Quantity"];
    if (colSales !== -1) {
      summaryHeaders.push("Total Sales");
    }
    
    newSheet.getRange(productBreakdownRow, productBreakdownStartCol, 1, summaryHeaders.length).setValues([summaryHeaders]);
    newSheet.getRange(productBreakdownRow, productBreakdownStartCol, 1, summaryHeaders.length).setFontWeight("bold");
    
    // Sort products by total sales (or quantity if no sales data)
    var productArray = Object.keys(productTotals).map(function(product) {
      return {
        name: product,
        qty: productTotals[product].qty,
        sales: productTotals[product].sales
      };
    });
    
    productArray.sort(function(a, b) {
      return colSales !== -1 ? b.sales - a.sales : b.qty - a.qty;
    });
    
    // Write product breakdown
    productBreakdownRow += 1;
    productArray.forEach(function(product, index) {
      var row = [product.name, product.qty];
      if (colSales !== -1) {
        row.push("$" + product.sales.toFixed(2));
      }
      
      newSheet.getRange(productBreakdownRow + index, productBreakdownStartCol, 1, row.length).setValues([row]);
    });
    
    // Auto-resize columns again to accommodate summary
    var maxCols = Math.max(headerRow.length, productBreakdownStartCol + summaryHeaders.length - 1);
    for (var col = 1; col <= maxCols; col++) {
      newSheet.autoResizeColumn(col);
    }
    
    Logger.log("Successfully created new sheet and wrote all data with auto-resized columns and summary");
    
  } catch (error) {
    Logger.log("Error with new sheet approach: " + error.toString());
    
    // Last resort - write a simple test to see what's happening
    Logger.log("Writing simple test data...");
    formatted.clear();
    
    // Write just a simple 3x3 grid to test
    var testData = [
      ["Location", "Product", "9/13"],
      ["Test Loc", "Test Product", 123],
      ["", "Test Product 2", 456]
    ];
    
    formatted.getRange(1, 1, 3, 3).setValues(testData);
    Logger.log("Test data written");
  }
}
