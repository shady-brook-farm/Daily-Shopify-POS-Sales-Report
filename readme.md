# Shopify POS Sales Data Processor for Google Sheets

A powerful Google Apps Script that transforms Shopify sales data exports into beautifully formatted pivot tables with advanced heatmap visualization and comprehensive analytics. Perfect for Shopify store owners and managers who need quick insights from their sales data.

<img width="1264" height="743" alt="image" src="https://github.com/user-attachments/assets/9096c855-8ed0-4504-a153-16c26feddcb5" />


## 🛍️ Why Use This Tool?

Shopify's native reporting can be overwhelming and doesn't always present data in the most actionable format. This script takes your raw Shopify sales exports and creates:

- **Clear visual patterns** in your sales data with color-coded heatmaps
- **Location-based comparisons** across multiple store locations
- **Product performance analysis** sorted by sales volume
- **Time-series visualization** showing daily sales trends
- **Professional formatting** ready for presentations and reports

## 🚀 Features

### Shopify-Specific Functionality
- **Automatic Shopify Data Processing**: Works seamlessly with Shopify POS and online sales exports
- **Multi-Location Support**: Perfect for businesses with multiple Shopify POS locations
- **Product Variant Analysis**: Handles both product titles and variant-specific data
- **Currency-Aware Calculations**: Properly formats sales totals and gross sales

### Advanced Analytics
- **Interactive Heatmap**: Color-coded cells showing sales patterns at a glance
- **Negative Value Detection**: Highlights returns and refunds in soft red-orange
- **Location Performance**: Compare sales across all your store locations
- **Product Ranking**: Automatic sorting by sales performance
- **Daily Trend Analysis**: Visualize sales patterns over time

### Professional Output
- **Executive-Ready Reports**: Clean formatting suitable for stakeholder presentations
- **Automated Processing**: Updates automatically when you paste new Shopify data
- **Visual Data Borders**: Professional table formatting
- **Summary Statistics**: Key metrics clearly displayed

## 📊 Perfect for Shopify Users Who Need:

- **Multi-location insights** from Shopify POS data
- **Product performance comparisons** across variants
- **Visual sales trend analysis** for inventory planning
- **Quick executive summaries** from raw Shopify exports
- **Return and refund analysis** (negative values highlighted)

## 🛠️ Getting Started with Your Shopify Data

### Step 1: Export Data from Shopify
Use this sample ShopifyQL query to export the perfect dataset:

#### ShopifyQL Query:
```sql
FROM sales
  SHOW net_items_sold, gross_sales, total_sales
  WHERE product_title IN ('PRODUCT NAMES HERE', 'COMMA SEPERATED')
  GROUP BY pos_location_name, day, product_title_at_time_of_sale,
    product_variant_title_at_time_of_sale WITH GROUP_TOTALS, TOTALS, CURRENCY 'USD'
  SINCE 2024-09-13 UNTIL 2024-10-30
  ORDER BY total_sales__pos_location_name_totals DESC, day ASC,
    total_sales__pos_location_name_day_product_title_at_time_of_sale_totals DESC,
    total_sales DESC, pos_location_name ASC, product_title_at_time_of_sale ASC,
    product_variant_title_at_time_of_sale ASC
VISUALIZE total_sales
```

**Note**: Replace `'PRODUCT NAMES HERE', 'COMMA SEPERATED'` with your actual product names, or remove the WHERE clause entirely to analyze all products.

### Step 2: Set Up Google Sheets
1. **Create a new Google Sheets** document
2. **Go to Extensions** → **Apps Script**
3. **Replace the default code** with the Sales Data Processor script. The code can be [found here](/code.gs)
4. **Save the project** as "Shopify Sales Processor"
5. **Return to your spreadsheet**

### Step 3: Import Your Shopify Data
1. **Create a sheet named "RawData"** (exact name required)
2. **Paste your Shopify export data** directly into this sheet
3. **The script will automatically process** your data when you make any edit
4. **View your formatted results** in the new "FormattedData" sheet

## 📋 Shopify Data Requirements

Your Shopify export must include these columns (standard Shopify export format):

| Required Shopify Columns | Optional Shopify Columns |
|-------------------------|-------------------------|
| `POS location name` | `Total sales` |
| `Day` | `Gross Sales` |
| `Net items sold` | `Product title at time of sale` |
| `Product variant title at time of sale` | |

### Sample Shopify Export Format
```csv
POS location name,Day,Product variant title at time of sale,Net items sold,Total sales,Gross Sales
Main Street Store,2024-01-15,Organic Coffee Beans - Dark Roast,12,36.00,40.00
Main Street Store,2024-01-15,Premium Tea Bags - Earl Grey,5,12.50,15.00
Downtown Location,2024-01-16,Organic Coffee Beans - Medium Roast,18,54.00,60.00
```

## 🎯 How to Use

### Automatic Processing (Recommended)
1. **Paste your Shopify export** into the "RawData" sheet
2. **Make any small edit** (like adding a space) to trigger processing
3. **Check the "FormattedData" sheet** for your beautiful results
4. **Refresh your data** anytime by pasting new exports and editing the sheet

### Manual Processing
If you prefer manual control:
1. **Open Extensions** → **Apps Script**
2. **Select `processCSVData`** from the function dropdown
3. **Click Run** to process your data
4. **View results** in the FormattedData sheet

## 📈 What You'll Get: Complete Shopify Analytics

### 1. Location & Product Pivot Table
Transform your raw Shopify data into a clean matrix showing:
- **Rows**: Your store locations and products
- **Columns**: Each day from your export
- **Values**: Color-coded sales quantities with heatmap
- **Visual cues**: Immediate pattern recognition

### 2. Executive Summary
Key metrics at a glance:
- **Total units sold** across all locations
- **Total revenue** (if sales data included)
- **Gross sales totals** (if gross sales data included)
- **Performance period** covered

### 3. Product Performance Ranking
Automatically sorted list showing:
- **Best performing products** by sales volume
- **Individual product totals** for quantities and revenue
- **Easy identification** of your top sellers
- **Underperforming items** that need attention

### 4. Visual Analytics
- **Positive sales**: White to deep blue gradient (higher = darker blue)
- **Returns/Refunds**: White to soft red-orange gradient (more returns = darker orange)
- **Zero activity**: Clean white background
- **Professional borders**: Clean table presentation

## 🏪 Shopify-Specific Benefits

### Multi-Location Insights
- **Compare performance** between your physical stores
- **Identify location-specific trends** for inventory planning
- **Spot seasonal patterns** across different markets
- **Optimize stock distribution** based on location performance

### Product & Variant Analysis
- **Understand variant preferences** (size, color, style)
- **Identify bestselling combinations** for marketing focus
- **Spot underperforming variants** for discontinuation
- **Plan inventory** based on variant-specific demand

### Return & Refund Tracking
- **Visualize negative values** (returns/refunds) in red-orange
- **Identify problematic products** with high return rates
- **Track refund patterns** by location or time period
- **Improve product quality** based on return data

## ⚙️ Customization for Your Shopify Store

### Adjust Date Ranges in ShopifyQL
Modify the SINCE and UNTIL dates in your query:
```sql
SINCE 2024-09-01 UNTIL 2024-12-31  -- Analyze Q4 performance
```

### Focus on Specific Products
Use the WHERE clause to analyze particular product lines:
```sql
WHERE product_title IN ('Winter Collection', 'Holiday Specials', 'New Arrivals')
```

### Change Color Intensity
In the script, modify color contrast for better visibility:
```javascript
var enhancedRatio = Math.pow(ratio, 0.5); // Current setting
// Use 0.3 for subtle colors, 0.8 for dramatic contrast
```

## 🚨 Shopify-Specific Troubleshooting

### "Missing required headers" Error
- **Check your ShopifyQL query** includes all required fields
- **Verify column names** match exactly (Shopify's standard format)
- **Ensure data export** completed successfully from Shopify

### No Data Showing
- **Verify your date range** in the ShopifyQL query captured sales
- **Check if product filter** is too restrictive
- **Confirm POS location names** are not empty in your export

### Colors Not Displaying
- **Ensure Net items sold** column contains numeric values (not text)
- **Check for currency symbols** in sales columns (should be numbers only)
- **Verify data isn't filtered** in Shopify before export

### Performance with Large Shopify Datasets
- **Limit date ranges** to 1-3 months for faster processing
- **Use product filters** in ShopifyQL to focus on specific categories
- **Process location by location** if you have many stores

## 💡 Tips for Shopify Users

### Best Practices
1. **Run weekly reports** to track trends and spot issues early
2. **Compare month-over-month** by processing different date ranges
3. **Share formatted results** directly with team members (no raw data confusion)
4. **Use for inventory planning** by identifying fast-moving products
5. **Track seasonal patterns** by comparing same periods across different years

### Advanced ShopifyQL Queries
**Focus on high-value products:**
```sql
WHERE total_sales > 1000
```

**Analyze specific time periods:**
```sql
SINCE -30d  -- Last 30 days
SINCE -1y UNTIL -1y+3m  -- Same quarter last year
```

**Include specific locations:**
```sql
WHERE pos_location_name IN ('Main Store', 'Mall Location')
```

## 📊 Sample Results

Transform this raw Shopify export:
```
Main Store,2024-01-15,Coffee Beans,12,$36.00,$40.00
Main Store,2024-01-16,Coffee Beans,8,$24.00,$28.00
Mall Store,2024-01-15,Coffee Beans,22,$66.00,$72.00
```

Into this beautiful analysis:
```
Location     Product        1/15    1/16    1/17
Main Store   Coffee Beans    12      8       15
            Tea Selection    5      12       7
Mall Store   Coffee Beans    22     18      25
            Pastries        -2     15      11
```

With color-coded visualization and complete analytics below!

## 🤝 Perfect for Shopify Teams

This tool is ideal for:
- **Store Managers**: Daily sales tracking and location comparison
- **Inventory Managers**: Product performance and restocking decisions  
- **Marketing Teams**: Identifying bestsellers for promotional focus
- **Executives**: High-level sales summaries and trend analysis
- **Multi-location Owners**: Comparative performance across stores

## 📄 Requirements

- **Shopify store** with POS or online sales data
- **Google Sheets** access (free Google account)
- **Basic familiarity** with Shopify's export functionality
- **ShopifyQL access** for advanced data queries

## 🎉 Get Started Today!

Ready to transform your Shopify sales data into actionable insights? 

1. **Copy the script** into Google Apps Script
2. **Export your Shopify data** using the provided ShopifyQL query
3. **Paste into the RawData sheet** and watch the magic happen
4. **Share beautiful reports** with your team in minutes

**Stop struggling with raw CSV exports and start making data-driven decisions!** 🚀📊

---

*Built specifically for Shopify merchants who want better insights from their sales data.*
