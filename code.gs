function onEdit(e) {
  if (e.source.getActiveSheet().getName() === "RawData") {
    processCSVData();
  }
}

function processCSVData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName("RawData");

  // Delete old FormattedData sheet if it exists
  if (ss.getSheetByName("FormattedData")) ss.deleteSheet(ss.getSheetByName("FormattedData"));
  var newSheet = ss.insertSheet("FormattedData");

  var data = raw.getDataRange().getValues();
  if (!data || data.length < 2) return;

  var headers = data[0].map(h => String(h).trim());
  var rows = data.slice(1);

  var colLocation   = headers.indexOf("POS location name");
  var colDate       = headers.indexOf("Day");
  var colVariant    = headers.indexOf("Product variant title at time of sale");
  var colProduct    = headers.indexOf("Product title at time of sale");
  var colQty        = headers.indexOf("Net items sold");
  var colGrossSales = headers.indexOf("Gross Sales");
  var colSales      = headers.indexOf("Total sales");

  if (colLocation === -1 || colDate === -1 || colQty === -1) {
    throw new Error("Missing required headers: POS location name, Day, Net items sold.");
  }

  var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  var pivot = {}; 
  var productTotals = {}; 
  var datesArray = [];
  var seenDates = new Set();
  var totalGrossSales = 0;

  rows.forEach(row => {
    var location = String(row[colLocation] || "").trim();
    if (!location) return;

    var dateObj = new Date(row[colDate]);
    if (!(dateObj instanceof Date) || isNaN(dateObj)) return;
    var dateLabel = Utilities.formatDate(dateObj, tz, "M/d");

    if (!seenDates.has(dateLabel)) {
      seenDates.add(dateLabel);
      datesArray.push({original: dateObj, formatted: dateLabel});
    }

    var product = String(row[colVariant] || row[colProduct] || "Unknown").trim();
    var qty = Number(row[colQty]) || 0;
    var sales = colSales !== -1 ? (Number(row[colSales]) || 0) : 0;
    var gross = colGrossSales !== -1 ? (Number(row[colGrossSales]) || 0) : 0;

    if (!pivot[location]) pivot[location] = {};
    if (!pivot[location][product]) pivot[location][product] = {};
    pivot[location][product][dateLabel] = (pivot[location][product][dateLabel] || 0) + qty;

    if (!productTotals[product]) productTotals[product] = { qty: 0, sales: 0, gross: 0 };
    productTotals[product].qty += qty;
    productTotals[product].sales += sales;
    productTotals[product].gross += gross;
    totalGrossSales += gross;
  });

  // Sort dates
  var dates = datesArray.sort((a,b)=>a.original-b.original).map(d=>d.formatted);

  // Build pivot output
  var output = [];
  var headerRow = ["Location","Product"].concat(dates);
  output.push(headerRow);

  Object.keys(pivot).forEach(location => {
    var firstProduct = true;
    // Sort products alphabetically
    Object.keys(pivot[location]).sort().forEach(product => {
      var rowOut = [];
      rowOut.push(firstProduct ? location : "");
      rowOut.push(product);
      dates.forEach(d => rowOut.push(pivot[location][product][d] || 0));
      output.push(rowOut);
      firstProduct = false;
    });
    output.push(Array(headerRow.length).fill("")); // spacer row
  });


  for (var i = 1; i < output.length; i++) {
    while (output[i].length < headerRow.length) output[i].push("");
    if (output[i].length > headerRow.length) output[i] = output[i].slice(0, headerRow.length);
  }

  // --- WRITE PIVOT TABLE ---
  newSheet.getRange(1,1,1,headerRow.length)
          .setValues([headerRow])
          .setFontWeight("bold")
          .setBackground("#d9d9d9");

  if (output.length > 1) {
    newSheet.getRange(2,1,output.length-1,headerRow.length)
            .setValues(output.slice(1));
  }

  for (var col=1; col<=headerRow.length; col++) newSheet.autoResizeColumn(col);

  // --- SUMMARY ---
  var summaryRow = output.length + 3;
  var summaryCol = 1;
  newSheet.getRange(summaryRow, summaryCol).setValue("SUMMARY")
          .setFontWeight("bold").setFontSize(14).setBackground("#d9d9d9");

  summaryRow += 2;
  var totalQty = Object.values(productTotals).reduce((a,b)=>a+b.qty,0);
  newSheet.getRange(summaryRow, summaryCol).setValue("Total Quantity Sold:").setFontWeight("bold");
  newSheet.getRange(summaryRow, summaryCol+1).setValue(totalQty);

  if (colSales !== -1) {
    var totalSales = Object.values(productTotals).reduce((a,b)=>a+b.sales,0);
    summaryRow += 1;
    newSheet.getRange(summaryRow, summaryCol).setValue("Total Sales:").setFontWeight("bold");
    newSheet.getRange(summaryRow, summaryCol+1).setValue("$"+totalSales.toFixed(2));
  }

  if (colGrossSales !== -1) {
    summaryRow += 1;
    newSheet.getRange(summaryRow, summaryCol).setValue("Total Gross Sales:").setFontWeight("bold");
    newSheet.getRange(summaryRow, summaryCol+1).setValue("$"+totalGrossSales.toFixed(2));
  }

  // --- PRODUCT BREAKDOWN BELOW SUMMARY ---
  var breakdownStartRow = summaryRow + 3;
  var breakdownCol = 1;
  newSheet.getRange(breakdownStartRow, breakdownCol).setValue("PRODUCT BREAKDOWN")
          .setFontWeight("bold").setFontSize(12).setBackground("#d9d9d9");
  breakdownStartRow += 2;

  var summaryHeaders = ["Product","Total Quantity"];
  if (colSales !== -1) summaryHeaders.push("Total Sales");
  if (colGrossSales !== -1) summaryHeaders.push("Gross Sales");

  newSheet.getRange(breakdownStartRow, breakdownCol,1,summaryHeaders.length)
          .setValues([summaryHeaders])
          .setFontWeight("bold")
          .setBackground("#d9d9d9");

  var productArray = Object.keys(productTotals).map(p => ({
    name: p,
    qty: productTotals[p].qty,
    sales: productTotals[p].sales,
    gross: productTotals[p].gross
  }));

  productArray.sort((a,b)=>colSales!==-1 ? b.sales-a.sales : b.qty-a.qty);

  productArray.forEach((product,index)=>{
    var row = [product.name, product.qty];
    if (colSales !== -1) row.push("$"+product.sales.toFixed(2));
    if (colGrossSales !== -1) row.push("$"+product.gross.toFixed(2));
    newSheet.getRange(breakdownStartRow+1+index, breakdownCol,1,row.length).setValues([row]);
  });

  var maxCols = Math.max(headerRow.length, summaryHeaders.length);
  for (var col=1; col<=maxCols; col++) newSheet.autoResizeColumn(col);
}
