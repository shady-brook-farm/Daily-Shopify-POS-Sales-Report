/**
 * Google Apps Script for Sales Data Processing and Visualization
 * 
 * This script automatically processes CSV data from a "RawData" sheet and creates
 * a formatted pivot table with heatmap visualization and summary statistics.
 * 
 * Features:
 * - Automatic pivot table generation by location, product, and date
 * - Color-coded heatmap for sales quantities
 * - Summary statistics and product breakdown
 * - Automatic formatting with borders and styling
 * 
 * @author Cameron Golden
 * @version 1.2
 * @since 8-20-2025
 */

/**
 * Triggers when a cell is edited in the spreadsheet
 * Automatically processes data when changes are made to the RawData sheet
 * 
 * @param {Object} e - The edit event object containing information about the edit
 */
function onEdit(e) {
  // Only process if the edit occurred in the RawData sheet
  if (e.source.getActiveSheet().getName() === "RawData") {
    processCSVData();
  }
}

/**
 * Interpolates color values for heatmap visualization
 * Creates a gradient from white to blue for positive values
 * Uses soft red-orange for negative values
 * Enhanced with exponential scaling for better color differentiation
 * 
 * @param {number} value - The current value to color
 * @param {number} maxValue - The maximum positive value in the dataset
 * @param {number} minValue - The minimum negative value in the dataset (optional)
 * @returns {string} RGB color string in format "rgb(r,g,b)"
 */
function interpolateColor(value, maxValue, minValue) {
  // Return white for zero values
  if (value === 0) return "#ffffff";
  
  // Handle negative values with soft red-orange coloring
  if (value < 0) {
    if (!minValue || minValue === 0) return "#ffcccc"; // Light red fallback
    
    // Calculate ratio for negative values (minValue should be negative)
    var negativeRatio = Math.abs(value) / Math.abs(minValue);
    var enhancedNegRatio = Math.pow(negativeRatio, 0.5);
    
    // Soft red-orange gradient from white (255,255,255) to red-orange (255,140,100)
    var r = Math.floor(255 - (255 - 255) * enhancedNegRatio);  // Stay at 255 (red component)
    var g = Math.floor(255 - (255 - 140) * enhancedNegRatio);  // 255 -> 140
    var b = Math.floor(255 - (255 - 100) * enhancedNegRatio);  // 255 -> 100
    
    return "rgb(" + r + "," + g + "," + b + ")";
  }
  
  // Handle positive values with blue coloring
  if (maxValue === 0) return "#ffffff";
  
  // Calculate ratio with exponential scaling for better color differentiation
  var ratio = value / maxValue;
  // Apply power function to increase color contrast (values between 0.3 and 2.0 work well)
  var enhancedRatio = Math.pow(ratio, 0.5); // Square root for more gradual transition
  
  // Enhanced color interpolation from white (255,255,255) to deep blue (100,150,200)
  var r = Math.floor(255 - (255 - 100) * enhancedRatio);  // 255 -> 100
  var g = Math.floor(255 - (255 - 150) * enhancedRatio);  // 255 -> 150
  var b = Math.floor(255 - (255 - 200) * enhancedRatio);  // 255 -> 200

  return "rgb(" + r + "," + g + "," + b + ")";
}

/**
 * Main function to process CSV data and create formatted output
 * 
 * This function:
 * 1. Reads data from the RawData sheet
 * 2. Creates a pivot table grouped by location, product, and date
 * 3. Applies heatmap coloring to visualize sales quantities
 * 4. Generates summary statistics and product breakdown
 * 5. Formats the output with proper styling and borders
 */
function processCSVData() {

  SpreadsheetApp.getActiveSpreadsheet().toast("Processing data...\n\n Please wait for success message.", "Status", 5);


  try {
    // Get reference to the active spreadsheet and raw data sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var raw = ss.getSheetByName("RawData");

    // Clean up: Delete existing FormattedData sheet if it exists
    var existingSheet = ss.getSheetByName("FormattedData");
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // Create new FormattedData sheet for output
    var newSheet = ss.insertSheet("FormattedData");

    // Get all data from the raw sheet
    var data = raw.getDataRange().getValues();
    
    // Validate data exists and has at least header + one data row
    if (!data || data.length < 2) {
      throw new Error("Insufficient data: Need at least header row and one data row");
    }

    // Process headers: trim whitespace and store
    var headers = data[0].map(function(h) { 
      return String(h).trim(); 
    });
    
    // Separate data rows from headers
    var rows = data.slice(1);

    // Map column indices for required fields
    var colLocation   = headers.indexOf("POS location name");
    var colDate       = headers.indexOf("Day");
    var colVariant    = headers.indexOf("Product variant title at time of sale");
    var colProduct    = headers.indexOf("Product title at time of sale");
    var colQty        = headers.indexOf("Net items sold");
    var colGrossSales = headers.indexOf("Gross Sales");
    var colSales      = headers.indexOf("Total sales");

    // Validate required columns exist
    if (colLocation === -1 || colDate === -1 || colQty === -1) {
      throw new Error("Missing required headers: POS location name, Day, Net items sold.");
    }

    // Get timezone for date formatting
    var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

    // Initialize data structures for processing
    var pivot = {};              // Main pivot table data structure
    var productTotals = {};      // Product-level totals for summary
    var datesArray = [];         // Array to store and sort dates
    var seenDates = new Set();   // Set to track unique dates
    var totalGrossSales = 0;     // Running total of gross sales

    // Process each data row
    rows.forEach(function(row) {
      // Extract and validate location
      var location = String(row[colLocation] || "").trim();
      if (!location) return; // Skip rows without location

      // Parse and validate date
      var dateObj = new Date(row[colDate]);
      if (!(dateObj instanceof Date) || isNaN(dateObj)) return; // Skip invalid dates
      
      // Format date for display (M/d format)
      var dateLabel = Utilities.formatDate(dateObj, tz, "M/d");

      // Track unique dates for column headers
      if (!seenDates.has(dateLabel)) {
        seenDates.add(dateLabel);
        datesArray.push({
          original: dateObj, 
          formatted: dateLabel
        });
      }

      // Extract product name (prefer variant over product title)
      var product = String(row[colVariant] || row[colProduct] || "Unknown").trim();
      
      // Parse numeric values with fallback to 0
      var qty = Number(row[colQty]) || 0;
      var sales = colSales !== -1 ? (Number(row[colSales]) || 0) : 0;
      var gross = colGrossSales !== -1 ? (Number(row[colGrossSales]) || 0) : 0;

      // Build pivot table structure: location -> product -> date -> quantity
      if (!pivot[location]) {
        pivot[location] = {};
      }
      if (!pivot[location][product]) {
        pivot[location][product] = {};
      }
      
      // Aggregate quantities by location/product/date
      pivot[location][product][dateLabel] = (pivot[location][product][dateLabel] || 0) + qty;

      // Build product totals for summary section
      if (!productTotals[product]) {
        productTotals[product] = { qty: 0, sales: 0, gross: 0 };
      }
      productTotals[product].qty += qty;
      productTotals[product].sales += sales;
      productTotals[product].gross += gross;
      
      // Track total gross sales
      totalGrossSales += gross;
    });

    // Sort dates chronologically for consistent column ordering
    var dates = datesArray
      .sort(function(a, b) { return a.original - b.original; })
      .map(function(d) { return d.formatted; });

    // Build pivot table output array
    var output = [];
    
    // Create header row: Location, Product, then date columns
    var headerRow = ["Location", "Product"].concat(dates);
    output.push(headerRow);

    // Process each location in the pivot data
    Object.keys(pivot).forEach(function(location) {
      var firstProduct = true; // Flag to show location name only once per location group
      
      // Sort products alphabetically within each location
      Object.keys(pivot[location]).sort().forEach(function(product) {
        var rowOut = [];
        
        // Add location name only for first product in each location group
        rowOut.push(firstProduct ? location : "");
        rowOut.push(product);
        
        // Add quantity data for each date column
        dates.forEach(function(d) {
          rowOut.push(pivot[location][product][d] || 0);
        });
        
        output.push(rowOut);
        firstProduct = false; // Subsequent rows in this location will have blank location cell
      });
      
      // Add spacer row between location groups for visual separation
      output.push(Array(headerRow.length).fill(""));
    });

    // Ensure all rows have consistent column count
    for (var i = 1; i < output.length; i++) {
      // Pad short rows with empty strings
      while (output[i].length < headerRow.length) {
        output[i].push("");
      }
      // Trim long rows to header length
      if (output[i].length > headerRow.length) {
        output[i] = output[i].slice(0, headerRow.length);
      }
    }

    // === WRITE PIVOT TABLE TO SHEET ===
    
    // Write and format header row
    newSheet.getRange(1, 1, 1, headerRow.length)
            .setValues([headerRow])
            .setFontWeight("bold")
            .setBackground("#d9d9d9");

    // Write data rows if they exist
    if (output.length > 1) {
      var dataRange = newSheet.getRange(2, 1, output.length - 1, headerRow.length);
      dataRange.setValues(output.slice(1));

      // === APPLY HEATMAP COLORS AND BORDERS ===
      for (var r = 2; r <= output.length; r++) {
        // Skip empty/spacer rows (rows where all cells are blank)
        var currentRow = output[r - 1];
        if (!currentRow || currentRow.every(function(cell) { return cell === ""; })) {
          continue;
        }
        
        // Get sales data values for this row (excluding Location and Product columns)
        var rowValues = newSheet.getRange(r, 3, 1, headerRow.length - 2).getValues()[0];
        
        // Find maximum positive and minimum negative values for color scaling
        var maxVal = Math.max.apply(Math, rowValues.filter(function(v) { return v > 0; }));
        var minVal = Math.min.apply(Math, rowValues.filter(function(v) { return v < 0; }));
        
        // Handle case where no positive or negative values exist
        if (!isFinite(maxVal)) maxVal = 0;
        if (!isFinite(minVal)) minVal = 0;
        
        // Generate color array for this row based on values
        var bgColors = rowValues.map(function(v) {
          return interpolateColor(Number(v), maxVal, minVal);
        });
        
        // Prepend white backgrounds for Location and Product columns
        var fullRowColors = ["#ffffff", "#ffffff"].concat(bgColors);
        
        // Apply background colors to entire row
        newSheet.getRange(r, 1, 1, headerRow.length).setBackgrounds([fullRowColors]);
        
        // Apply top and bottom borders only to sales data columns (3 onwards)
        var borderRange = newSheet.getRange(r, 3, 1, headerRow.length - 2);
        borderRange.setBorder(
          true,  // top
          null,  // left
          true,  // bottom
          null,  // right
          null,  // vertical
          null,  // horizontal
          "#000000", // color
          SpreadsheetApp.BorderStyle.SOLID // style
        );
      }
    }

    // resize all columns for better readability
    for (var col = 1; col <= headerRow.length; col++) {
      newSheet.autoResizeColumn(col);
    }

    // Create Summary Section
    
    var summaryRow = output.length + 3; // Start summary below pivot table with spacing
    var summaryCol = 1;
    
    // Summary section header
    newSheet.getRange(summaryRow, summaryCol)
            .setValue("SUMMARY")
            .setFontWeight("bold")
            .setFontSize(14)
            .setBackground("#d9d9d9");

    summaryRow += 2; // Add spacing after header

    // Calculate and display total quantity sold
    var totalQty = Object.values(productTotals).reduce(function(acc, product) {
      return acc + product.qty;
    }, 0);
    
    newSheet.getRange(summaryRow, summaryCol)
            .setValue("Total Quantity Sold:")
            .setFontWeight("bold");
    newSheet.getRange(summaryRow, summaryCol + 1)
            .setValue(totalQty);

    // Add total sales if sales column exists
    if (colSales !== -1) {
      var totalSales = Object.values(productTotals).reduce(function(acc, product) {
        return acc + product.sales;
      }, 0);
      
      summaryRow += 1;
      newSheet.getRange(summaryRow, summaryCol)
              .setValue("Total Sales:")
              .setFontWeight("bold");
      newSheet.getRange(summaryRow, summaryCol + 1)
              .setValue("$" + totalSales.toFixed(2));
    }

    // Add gross sales if gross sales column exists
    if (colGrossSales !== -1) {
      summaryRow += 1;
      newSheet.getRange(summaryRow, summaryCol)
              .setValue("Total Gross Sales:")
              .setFontWeight("bold");
      newSheet.getRange(summaryRow, summaryCol + 1)
              .setValue("$" + totalGrossSales.toFixed(2));
    }

    // Create product breakdown section
    
    var breakdownStartRow = summaryRow + 3; // Start below summary with spacing
    var breakdownCol = 1;
    
    // Product breakdown header
    newSheet.getRange(breakdownStartRow, breakdownCol)
            .setValue("PRODUCT BREAKDOWN")
            .setFontWeight("bold")
            .setFontSize(12)
            .setBackground("#d9d9d9");
    
    breakdownStartRow += 2; // Add spacing after header

    // Build headers for product breakdown table
    var summaryHeaders = ["Product", "Total Quantity"];
    if (colSales !== -1) summaryHeaders.push("Total Sales");
    if (colGrossSales !== -1) summaryHeaders.push("Gross Sales");

    // Write and format breakdown table headers
    newSheet.getRange(breakdownStartRow, breakdownCol, 1, summaryHeaders.length)
            .setValues([summaryHeaders])
            .setFontWeight("bold")
            .setBackground("#d9d9d9");

    // Convert product totals to sortable array
    var productArray = Object.keys(productTotals).map(function(productName) {
      return {
        name: productName,
        qty: productTotals[productName].qty,
        sales: productTotals[productName].sales,
        gross: productTotals[productName].gross
      };
    });

    // Sort products by sales (if available) or quantity (descending order)
    productArray.sort(function(a, b) {
      return colSales !== -1 ? b.sales - a.sales : b.qty - a.qty;
    });

    // Write product breakdown data
    productArray.forEach(function(product, index) {
      var row = [product.name, product.qty];
      
      // Add sales data if columns exist
      if (colSales !== -1) row.push("$" + product.sales.toFixed(2));
      if (colGrossSales !== -1) row.push("$" + product.gross.toFixed(2));
      
      // Write row to sheet
      newSheet.getRange(breakdownStartRow + 1 + index, breakdownCol, 1, row.length)
              .setValues([row]);
    });

    // Final auto-resize for all columns to accommodate all content
    var maxCols = Math.max(headerRow.length, summaryHeaders.length);
    for (var col = 1; col <= maxCols; col++) {
      newSheet.autoResizeColumn(col);
    }

  } catch (error) {
    // Error handling: Display user-friendly error message
    console.error("Error processing CSV data: " + error.message);
    SpreadsheetApp.getUi().alert(
      "Error Processing Data: " + error.message
    );
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Automation Complete!", "Success!", 3);
}
