# Copilot with SQL and Excel — Assignment Answer Sheet

This repository contains the completed answers for the **"Copilot with SQL and Excel"** assignment (PW Skills), covering GitHub Copilot in SSMS (Part A) and Excel Copilot (Part B).

## 📁 Repository Structure

```
Copilot-SQL-Excel-Assignment/
├── README.md                          # This file — full answer sheet
├── sql/
│   ├── Q1_create_table_insert_data.sql
│   ├── Q2_sales_per_city_top5.sql
│   └── Q3_above_average_customers.sql
├── vba/
│   └── Q10_DataCleaning.bas
├── excel/
│   ├── Excel_dataset_cleaned.csv      # Cleaned dataset (Q4 + Q5 applied)
│   └── pivot_region_product.csv       # Pivot output (Q7)
└── images/
    └── region_sales_chart.png         # Chart for Q6
```

---

# Part A — Copilot with SQL (SSMS)

> Tool: GitHub Copilot inside SQL Server Management Studio (SSMS). Full scripts are in the `/sql` folder.

## Question 1 — Create `SalesData` table + insert 10,000 rows

**Copilot prompt used:** *"Create a table SalesData with columns CustomerID, Name, Age, City, PurchaseAmount, PurchaseDate, then insert 10,000 rows of random data."*

Copilot generates the `CREATE TABLE` statement, then — because a row-by-row loop for 10,000 inserts is slow — suggests a **numbers/tally CTE** (cross join on `sys.all_objects`) combined with `NEWID()`/`CHECKSUM()` to generate randomized Name, Age, City, PurchaseAmount and PurchaseDate values in a single set-based `INSERT ... SELECT`.

See: [`sql/Q1_create_table_insert_data.sql`](sql/Q1_create_table_insert_data.sql)

```sql
CREATE TABLE SalesData (
    CustomerID    INT IDENTITY(1,1) PRIMARY KEY,
    Name          VARCHAR(50),
    Age           INT,
    City          VARCHAR(50),
    PurchaseAmount DECIMAL(10,2),
    PurchaseDate  DATE
);
-- + set-based INSERT of 10,000 randomized rows (full script in /sql)
```

**Result:** `SELECT COUNT(*) FROM SalesData;` → `10000`

---

## Question 2 — Total sales per city & Top 5 cities by revenue

**Copilot prompts used:** *"Find total sales per city"* and *"Top 5 cities by revenue"*

```sql
-- Total sales per city
SELECT City, SUM(PurchaseAmount) AS TotalSales,
       COUNT(*) AS NumberOfOrders, AVG(PurchaseAmount) AS AvgOrderValue
FROM SalesData
GROUP BY City
ORDER BY TotalSales DESC;

-- Top 5 cities by revenue
SELECT TOP 5 City, SUM(PurchaseAmount) AS TotalRevenue
FROM SalesData
GROUP BY City
ORDER BY TotalRevenue DESC;
```

Full script: [`sql/Q2_sales_per_city_top5.sql`](sql/Q2_sales_per_city_top5.sql)

Because the data is randomly generated (Question 1), the exact top-5 city names will vary each run — but the query pattern above is the correct, Copilot-verified way to answer it: `GROUP BY City` + `SUM(PurchaseAmount)` + `ORDER BY ... DESC` + `TOP 5`.

---

## Question 3 — Customers with purchases above average

**Copilot prompt used:** *"Find customers with purchases above average"*

```sql
SELECT CustomerID, Name, City, PurchaseAmount,
       (SELECT AVG(PurchaseAmount) FROM SalesData) AS OverallAverage
FROM SalesData
WHERE PurchaseAmount > (SELECT AVG(PurchaseAmount) FROM SalesData)
ORDER BY PurchaseAmount DESC;
```

Copilot uses a **scalar subquery** to compute `AVG(PurchaseAmount)` and filters the outer query against it — the standard pattern for "above average" queries. A second query (in the script) also reports what percentage of customers fall above average.

Full script: [`sql/Q3_above_average_customers.sql`](sql/Q3_above_average_customers.sql)

---

# Part B — Copilot in Excel

> Tool: Excel Copilot, applied to `Dummy_dataset` (`Excel_dataset.csv` — 550 rows: `Customer_Name, Region, Product, Sales, Quantity, Order_Date`).

## Question 4 — Identify and handle missing values

**Copilot prompt used:** *"Identify missing values in this dataset and suggest how to handle them."*

Missing-value scan of the raw dataset:

| Column | Missing Values |
|---|---|
| Region | 192 |
| Product | 105 |
| Sales | 167 |
| Quantity | 90 |
| Customer_Name | 68 |
| Order_Date | 0 |

**Copilot's suggested handling (applied):**
- **Region / Product / Customer_Name** (text columns) → filled with `"Unknown"` (safer than dropping ~35% of rows).
- **Sales / Quantity** (numeric columns) → filled with the **column median**, which is more robust to outliers than the mean.
- **Quantity** also contained text entries like `"two"` instead of `2` — these were converted to numbers before imputation.

Cleaned output: [`excel/Excel_dataset_cleaned.csv`](excel/Excel_dataset_cleaned.csv)

---

## Question 5 — Identify & remove duplicate rows

**Copilot prompt used:** *"Find duplicate rows in this table and remove them."*

- Duplicate rows found (exact full-row matches): **51 rows**
- Copilot's suggestion: use **Data → Remove Duplicates** (or the `=COUNTIFS()` helper-column trick to flag duplicates first, then delete), checking all columns.
- Rows after removing duplicates: `550 → 499`

---

## Question 6 — Chart: which region has the highest sales

**Copilot prompt used:** *"Create a column chart of total sales by region and tell me which region has the highest sales."*

![Total Sales by Region](images/region_sales_chart.png)

**Interpretation:**
Excluding the `Unknown` region bucket (rows where Region was originally missing — a data-quality artifact, not a real region), **West has the highest total sales (₹56,900)**, followed by South (₹51,000), East (₹47,800), and North (₹42,300). This means the West region is the strongest-performing territory in the cleaned dataset, and would be the natural focus for a "top region" business insight.

---

## Question 7 — Pivot table: total sales by region and product

**Copilot prompt used:** *"Create a pivot table showing total sales by region and product."*

| Region | Laptop | Monitor | Phone | Tablet | Unknown |
|---|---|---|---|---|---|
| East | 6,400 | 13,300 | 9,200 | 8,400 | 10,500 |
| North | 7,300 | 10,600 | 5,900 | 7,600 | 10,900 |
| South | 12,500 | 9,800 | 9,200 | 9,400 | 10,100 |
| West | 13,300 | 11,700 | 17,900 | 8,100 | 5,900 |
| Unknown | 21,800 | 26,800 | 21,700 | 24,700 | 23,800 |

Full CSV export: [`excel/pivot_region_product.csv`](excel/pivot_region_product.csv)

**Reading the pivot table:** West sells the most Phones (₹17,900) of any region/product pair among known regions; South leads on Laptops (₹12,500). The `Unknown` row/column reflects rows with missing Region or Product data (Q4) rather than a genuine category.

---

## Question 8 — Calculated column: Total Sales

**Copilot prompt used:** *"Add a calculated column Total Sales = Sales × Quantity."*

Excel formula Copilot inserts (row 2 shown, filled down the table):

```
=[@Sales]*[@Quantity]
```
or in classic A1 notation: `=D2*E2`

This column is included as `Total_Sales` in [`excel/Excel_dataset_cleaned.csv`](excel/Excel_dataset_cleaned.csv).

---

## Question 9 — Dataset summary & key insights

**Copilot prompt used:** *"Summarize this dataset and give me key insights."*

- **Total Sales (cleaned dataset):** ₹316,800 across 499 orders
- **Average order value:** ≈ ₹635
- **Best-performing known region:** West (₹56,900)
- **Product ranking by revenue:** Monitor (₹72,200) > Phone (₹63,900) > Laptop (₹61,300) > Tablet (₹58,200)
- **Top customers by total spend:** Sneha (₹54,000), Amit (₹51,900), Karan (₹44,600)
- **Data quality note:** ~20–35% of fields across Region/Product/Sales/Quantity/Customer_Name were originally missing, and Quantity contained mixed text/numeric entries (e.g. `"two"`). This should be flagged to whoever owns the data source, since it materially affects region/product-level conclusions (the `Unknown` bucket is currently the single largest "region" by volume simply because of missing data, not real sales).

---

## Question 10 — VBA code for data cleaning

**Copilot prompt used:** *"Generate VBA code for cleaning this data — remove duplicates, fill null values, and format as a table."*

Full macro: [`vba/Q10_DataCleaning.bas`](vba/Q10_DataCleaning.bas)

The `CleanData()` macro:
1. Scans each column — fills blank text cells with `"Unknown"` and blank numeric cells with the column **median**.
2. Removes duplicate rows across all columns (`Range.RemoveDuplicates`).
3. Converts the cleaned range into a formatted Excel Table (`ListObject`, style `TableStyleMedium9`) and autofits columns.

```vba
Sub CleanData()
    ' ... see vba/Q10_DataCleaning.bas for the full macro
End Sub
```

To use: open the workbook → `Alt+F11` → Insert → Module → paste the code → run `CleanData`.

---

## How to reproduce

1. **SQL part:** open SSMS → connect to your instance → run the scripts in `/sql` in order (Q1 → Q2 → Q3), with GitHub Copilot enabled to see live suggestions.
2. **Excel part:** open `Dummy_dataset` in Excel with Copilot enabled → follow the prompts listed above → compare against `excel/Excel_dataset_cleaned.csv` and `excel/pivot_region_product.csv`.
