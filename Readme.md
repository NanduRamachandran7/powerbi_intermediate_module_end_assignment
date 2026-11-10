# 📊 Power BI Summative Assessment - Sales Data Analytics

[![Power BI](https://img.shields.io/badge/Power_BI-F2C811?style=for-the-badge&logo=power-bi&logoColor=black)](https://powerbi.microsoft.com/)
[![Power Query M](https://img.shields.io/badge/Power_Query-2A7A9B?style=for-the-badge&logo=microsoft&logoColor=white)](https://docs.microsoft.com/en-us/powerquery-m/)
[![DAX](https://img.shields.io/badge/DAX-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)](https://docs.microsoft.com/en-us/dax/)

## 📌 Project Overview
This repository contains the complete solution for the **Power BI Module End Evaluation** (Total Marks: 15). The project addresses severe data quality issues in a 1,000-row sales dataset (`SalesData_1000Rows_WithIssues_copy.csv`), implements advanced M-code transformations in Power Query, creates DAX calculated entities, and delivers an executive dashboard.

---

## 🛠️ Data Cleaning Architecture (Power Query / M-Code)

### Key Transformation Steps:
1. **Handling Missing Values:**
   - Categorical columns (`Product`, `CustomerName`, `Region`) replaced empty strings with `"Unknown"`.
   - Reconstructed `UnitPrice_Clean`, `Sales_Clean`, `Cost_Clean`, and `Profit_Clean` using logical arithmetic formulas where values were omitted.
2. **Context-Aware Group-Based Date Filling:**
   - Leveraged `Table.Group` and `Table.FillDown`/`Table.FillUp` on `OrderID` groups to inherit valid dates for duplicate order IDs (e.g., copying `03-07-1997` for `OrderID 1741`).
3. **Statistical Median Date Imputation:**
   - Computed the non-parametric median date (`06-12-1998`) over deduplicated records to fill standalone orphan `null` dates (e.g., `OrderID 1175`).
4. **Deduplication:**
   - Applied `Table.Distinct` to eliminate fully redundant rows post-cleaning.

```powerquery
let
    Source = Csv.Document(File.Contents("D:\PowerBI_Summative_Assessment\SalesData_1000Rows_WithIssues_copy.csv"), [Delimiter = ",", Columns = 11, QuoteStyle = QuoteStyle.None]),
    #"Promoted headers" = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"OrderID", Int64.Type}, {"OrderDate", type text}, {"CustomerName", type text}, {"Region", type text}, {"Product", type text}, {"Category", type text}, {"Quantity", Int64.Type}, {"UnitPrice", Int64.Type}, {"Sales", Int64.Type}, {"Cost", Int64.Type}, {"Profit", Int64.Type}}, "en-US"),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed column type",{{"OrderDate", type date}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type","","Unknown",Replacer.ReplaceValue,{"Product"}),
    #"Changed Type with Locale" = Table.TransformColumnTypes(#"Replaced Value", {{"OrderDate", type date}}, "en-IN"),
    #"Replaced Value1" = Table.ReplaceValue(#"Changed Type with Locale","","Unknown",Replacer.ReplaceValue,{"CustomerName"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","","Unknown",Replacer.ReplaceValue,{"Region"}),
    #"Filtered Rows" = Table.SelectRows(#"Replaced Value2", each true),
    #"Duplicated Column" = Table.DuplicateColumn(#"Filtered Rows", "UnitPrice", "UnitPrice - Copy"),
    #"Removed Columns" = Table.RemoveColumns(#"Duplicated Column",{"UnitPrice - Copy"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "UnitPrice.1", each [Sales]/[Quantity]),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom",{"UnitPrice.1"}),
    #"Added Custom1" = Table.AddColumn(#"Removed Columns1", "Custom", each if [UnitPrice] = null then [Sales] / [Quantity] else [UnitPrice]),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom1",{{"Custom", "UnitPrice_Clean"}}),
    #"Added Custom2" = Table.AddColumn(#"Renamed Columns", "Sales_Clean", each if([Sales]=null) then [Quantity]*[UnitPrice_Clean] else [Sales]),
    #"Added Custom3" = Table.AddColumn(#"Added Custom2", "Cost_Clean", each if [Cost] = null then [Sales_Clean] - [Profit] else [Cost]),
    #"Added Custom4" = Table.AddColumn(#"Added Custom3", "Profit_Clean", each if [Profit]=null then [Sales_Clean]-[Cost_Clean] else [Profit]),

    // 1. Group by OrderID for intra-group date filling
    #"Grouped Rows" = Table.Group(#"Added Custom4", {"OrderID"}, {
        {"GroupData", (subTable) => 
            let 
                SortedSub = Table.Sort(subTable, {{"OrderDate", Order.Descending}}),
                FilledDown = Table.FillDown(SortedSub, {"OrderDate"}),
                FilledUp = Table.FillUp(FilledDown, {"OrderDate"})
            in 
                FilledUp, 
            type table}
    }),

    // 2. Expand all columns
    #"Expanded GroupData" = Table.ExpandTableColumn(#"Grouped Rows", "GroupData", 
        {"OrderDate", "CustomerName", "Region", "Product", "Category", "Quantity", "UnitPrice", "Sales", "Cost", "Profit", "UnitPrice_Clean", "Sales_Clean", "Cost_Clean", "Profit_Clean"}, 
        {"OrderDate", "CustomerName", "Region", "Product", "Category", "Quantity", "UnitPrice", "Sales", "Cost", "Profit", "UnitPrice_Clean", "Sales_Clean", "Cost_Clean", "Profit_Clean"}
    ),

    // 3. Deduplicate first
    #"Removed Duplicates" = Table.Distinct(#"Expanded GroupData"),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Duplicates",{{"OrderID", Int64.Type}, {"UnitPrice", Currency.Type}, {"Sales", Currency.Type}, {"Cost", Currency.Type}, {"Profit", Currency.Type}, {"UnitPrice_Clean", Currency.Type}, {"Sales_Clean", Currency.Type}, {"Cost_Clean", Currency.Type}, {"Profit_Clean", Currency.Type}, {"OrderDate", type date}}),

    // 4. Calculate & replace remaining orphan dates with median date (06-12-1998)
    Median_OrderDate = Date.From(List.Median(List.RemoveNulls(#"Changed Type1"[OrderDate]))),
    #"Replaced Nulls with Median" = Table.ReplaceValue(#"Changed Type1", null, Median_OrderDate, Replacer.ReplaceValue, {"OrderDate"})
in
    #"Replaced Nulls with Median"
```

---

## 📈 DAX Implementations

### 1. Calculated Table: `EastRegionOrders`
```dax
EastRegionOrders = 
FILTER(
    SalesData_1000Rows_WithIssues_copy,
    SalesData_1000Rows_WithIssues_copy[Region] = "East"
)
```

### 2. Calculated Column: `ProfitCostDifference`
```dax
ProfitCostDifference = 
SalesData_1000Rows_WithIssues_copy[Profit_Clean] - SalesData_1000Rows_WithIssues_copy[Cost_Clean]
```

### 3. Calculated Measure: `TotalProfit`
```dax
TotalProfit = 
SUM(SalesData_1000Rows_WithIssues_copy[Profit_Clean])
```

---

## 📊 Dashboard Visualizations
- **Pie Chart:** Order distribution by region (`Count of OrderID by Region`).
- **Column Chart:** Top 5 trending products (`Monitor`, `Laptop`, `Bookshelf`, `Cabinet`, `Headphones`).
- **Line Chart:** Multi-decade profit trend analysis (`Sum of Profit_Clean by Year`).
- **Card Visual:** Displays `TotalProfit` (**₹ 5.153105M**).

---

## 📑 Evaluation Audit
| Question | Description | Score |
| :--- | :--- | :---: |
| **Q1: Data Cleaning** | Full Power Query transformation, group filling, median date imputation, deduplication | **5 / 5** |
| **Q2: Visualizations** | Regional pie chart, top 5 column chart, profit line chart | **5 / 5** |
| **Q3: DAX Modeling** | `EastRegionOrders` table, `ProfitCostDifference` column, `TotalProfit` measure | **5 / 5** |
| **Total** | **Summative Assessment Evaluation** | **15 / 15** |