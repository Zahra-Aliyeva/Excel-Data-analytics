# 📊 Excel Data Analytics — Superstore Sales Project

An Excel analytics project using the Superstore retail dataset. I worked with raw data, cleaned and organized it, then used pivot tables, lookup formulas, calculated fields, and conditional formatting to analyze the data and build an interactive sales dashboard.

![Dashboard Preview](Screenshot.png)

## 🗂️ Dataset

| Sheet | Description |
|---|---|
| Orders | Order-level transactions: dates, customer, region, product, sales, quantity, discount, profit |
| Products | Product catalog with supplier, unit cost, and target margin |
| People | Regional managers mapped to each sales region |
| Returns | Order IDs that were returned |

## ✅ Checkpoint 1 — Data Cleaning
[Checkpoint-1.xlsx](Checkpoint-1 (1).xlsx)

- Removed duplicate rows using Excel's built-in duplicate removal (10 duplicates found and removed).
- Deleted a completely blank row (row 1872).
- Fixed inconsistent data types:
  - Order Date — was stored as Text, converted to Short Date.
  - Ship Date — already correctly formatted as Date.
  - Unit Cost, Sales, Quantity, Discount — were General, converted to Number.
- Standardized formatting across the sheet:
  - Unified font to Calibri, size 11 (previously mixed with Times New Roman and inconsistent sizes).
  - Applied AutoFit to all column widths.
  - Bolded all column headers.

## ✅ Checkpoint 2 — Pivot Tables
[Checkpoint-2.xlsx](Checkpoint-2.xlsx)

Five business questions answered using PivotTables:

**1. Which product sub-category generates the highest total sales?**
Sub-Category in Rows, Sum of Sales in Values, sorted descending → Chairs ranked highest, followed by Phones and Storage.

**2. Which product category generates the highest profit in each region?**
Segment as a filter, Category in Rows, Region in Columns, Sum of Profit in Values → Technology comes out on top in every single region.

**3. How do sales and profit compare across different shipping modes?**
Region as a filter, Category in Rows, Ship Mode in Columns, Sum of Sales and Sum of Profit in Values. Standard Class pulls ahead of every other shipping mode on both sales and profit — not surprising, since it's usually the default option customers don't bother changing.

**4. How have sales and profit changed across regions over the years?**
Order Date grouped by Year, Segment as a filter, Year in Rows, Region in Columns, Sum of Sales and Sum of Profit in Values. The results show something worth double-checking before this goes in front of anyone outside the team: total sales dropped from $1.82M in 2023 to $40.9K in 2026, while profit moved the opposite direction, climbing from $51.7K to $95.9K over the same four years. That swing is big enough that it's worth confirming with whoever maintains the Orders sheet — it could reflect a real shift toward fewer, higher-margin sales, or it could mean 2026 isn't a complete year of data yet.

**5. Which customer segment receives the highest average discount?**
Region as a filter, Segment in Rows, Average of Discount in Values. Consumer and Corporate both average a 16% discount, Home Office comes in slightly lower at 15%.

## ✅ Checkpoint 3 — Lookup Formulas
[Checkpoint-3.xlsb](Checkpoint-3.xlsb)

Combined data from People and Products into the Orders sheet using both INDEX-MATCH and XLOOKUP, plus a VLOOKUP to flag returns:

| New Column | Method | Formula |
|---|---|---|
| Regional Manager | INDEX-MATCH | `=INDEX(People!$A$2:$A$5,MATCH(M2,People!$B$2:$B$5,0))` |
| Regional Manager | XLOOKUP | `=XLOOKUP(M2,People!$B:$B,People!$A:$A,"Not Found")` |
| Supplier | INDEX-MATCH | `=INDEX(Products!$C:$C,MATCH(N2,Products!$A:$A,0))` |
| Unit Cost | INDEX-MATCH | `=INDEX(Products!$D:$D,MATCH(N2,Products!$A:$A,0))` |
| Returned | VLOOKUP | `=IFERROR(VLOOKUP(B2,Returns!$A:$B,2,FALSE),"No")` |
| Target Margin | XLOOKUP | `=XLOOKUP(N2,Products!$A:$A,Products!$E:$E,"Not Found")` |

Regional Manager was pulled two different ways on purpose — INDEX-MATCH and XLOOKUP side by side — just to confirm both land on the same result. They did.

## ✅ Checkpoint 4 — Calculated Fields
[Checkpoint-4.xlsb](Checkpoint-4.xlsb)

IF / Nested IF / IFS classifications:

| New Column | Logic | Formula |
|---|---|---|
| Profit Status | Profitable vs. Loss | `=IF(U2>0,"Profitable","Loss")` |
| Discount Level | No Discount / Low / Medium / High | `=IF(T2=0,"No Discount",IF(T2<=0.1,"Low",IF(T2<=0.2,"Medium","High")))` |
| Margin Category | Low / Medium / High Margin | `=IFS(AB2<0.2,"Low Margin",AB2<0.3,"Medium Margin",AB2>=0.3,"High Margin")` |

SUMIFS / COUNTIFS summary table (by region):

| Metric | Formula (example: West) |
|---|---|
| Total Sales by Region | `=SUMIFS(Orders!$R:$R,Orders!$M:$M,"West")` |
| Returned Orders by Region | `=COUNTIFS(Orders!M:M,"West",Orders!AA:AA,"Yes")` |

Same formulas repeated for East, Central, and South, just swapping the region criterion.

⚠️ Quality check: confirmed the SUMIFS/COUNTIFS criteria ranges and sum range were the same size — a mismatch here fails silently and just returns 0 instead of throwing an error.

## ✅ Checkpoint 5 — Interactive Dashboard
[Checkpoint-5.xlsb](Checkpoint-5.xlsb)

Built using PivotTables, PivotCharts, KPI cards, Slicers, and Conditional Formatting.

**KPI Cards**

| KPI | Value | Formula |
|---|---|---|
| Total Sales | $2.33M | `=SUM(Orders!R:R)` |
| Total Profit | $292.3K | `=SUM(Orders!U:U)` |
| Total Orders | 5,111 | `=COUNTA(UNIQUE(Orders!B2:B10195))` |
| Profit Margin % | 12.6% | `=SUM(Orders!U:U)/SUM(Orders!R:R)` |

**Charts (4 types)**

| Chart | Type | Insight |
|---|---|---|
| Total Sales by Region | Clustered Column | West leads at $671K, South trails at $507K |
| Top 5 Products by Sales | Horizontal Bar | Revenue is spread across products rather than concentrated in one — the top seller (a message book multi-pack) brings in $25.9K, barely ahead of staple envelopes and binding combs |
| Sales Trend by Year | Line | Total sales fell steadily, from $1.82M in 2023 down to $40.9K in 2026, while profit climbed the whole time, from $51.7K to $95.9K |
| Sales vs Profit by Category | Clustered Column | Office Supplies brings in far more revenue ($1.44M) than the other categories, but Technology actually returns more profit ($147K vs. $126K) on less than a third of the sales |

Slicers — Segment, Region, and Years — all connected to the PivotCharts for interactive filtering.

**Conditional Formatting**

- Profit → Green (profitable) / Red (loss)
- Unit Cost → Data Bars, longer bar = higher cost
- Discount Level → Green (No Discount), Blue (Low), Yellow (Medium), Red (High)

## 💡 Key Insights & Recommendations

A few things stood out while putting this together, worth flagging to the sales and finance teams:

**Technology is quietly the strongest category, not Office Supplies.** Office Supplies drives the most revenue on the dashboard, but Technology converts more of its sales into actual profit — $147K in profit from $386K in sales, against $126K from $1.44M. If budget or sales attention is being allocated by "which category sells the most," that's pointing at the wrong one.

**Furniture is barely paying for itself.** $497K in sales produced only $20K in profit, a margin under 5%. Before writing that off as normal for the category, it's worth checking whether Furniture is simply being discounted more heavily than everything else — even a small pullback there could move real money to the bottom line.

**The discount gap between segments is small but consistent.** Consumer and Corporate average a 16% discount, Home Office averages 15%. One point isn't much on its own, but if Home Office customers are ordering similar volumes without needing the extra discount, that points to the other two segments being over-discounted rather than Home Office being under-served.

**Standard Class shipping wins mostly because it's the default, not necessarily because it's the best option.** It's fine that it leads on volume, but profit-per-order by shipping mode is worth pulling separately — expedited options like Same Day or First Class often carry costs that a plain sales/profit comparison won't surface.

**No single product is carrying the business.** The top 5 products by sales all sit in a tight $14.6K–$25.9K range, so there's no one SKU the business is overly dependent on. That's healthy from a risk standpoint, but it also means the more realistic path to growth is pushing Technology harder alongside these steady sellers, not chasing one big hit product.

## ✅ Checkpoint 6 — Formula Documentation

This README doubles as the formula documentation, consolidating every formula used across the project:

| Purpose | Formula |
|---|---|
| Regional Manager (INDEX-MATCH) | `=INDEX(People!$A$2:$A$5,MATCH(M2,People!$B$2:$B$5,0))` |
| Regional Manager (XLOOKUP) | `=XLOOKUP(M2,People!$B:$B,People!$A:$A,"Not Found")` |
| Supplier (INDEX-MATCH) | `=INDEX(Products!$C:$C,MATCH(N2,Products!$A:$A,0))` |
| Unit Cost (INDEX-MATCH) | `=INDEX(Products!$D:$D,MATCH(N2,Products!$A:$A,0))` |
| Returned (VLOOKUP) | `=IFERROR(VLOOKUP(B2,Returns!$A:$B,2,FALSE),"No")` |
| Target Margin (XLOOKUP) | `=XLOOKUP(N2,Products!$A:$A,Products!$E:$E,"Not Found")` |
| Profit Status (IF) | `=IF(U2>0,"Profitable","Loss")` |
| Discount Level (Nested IF) | `=IF(T2=0,"No Discount",IF(T2<=0.1,"Low",IF(T2<=0.2,"Medium","High")))` |
| Margin Category (IFS) | `=IFS(AB2<0.2,"Low Margin",AB2<0.3,"Medium Margin",AB2>=0.3,"High Margin")` |
| Total Sales by Region (SUMIFS) | `=SUMIFS(Orders!$R:$R,Orders!$M:$M,"West")` |
| Returned Orders by Region (COUNTIFS) | `=COUNTIFS(Orders!M:M,"West",Orders!AA:AA,"Yes")` |
| Total Sales KPI | `=SUM(Orders!R:R)` |
| Total Profit KPI | `=SUM(Orders!U:U)` |
| Total Orders KPI | `=COUNTA(UNIQUE(Orders!B2:B10195))` |
| Profit Margin % KPI | `=SUM(Orders!U:U)/SUM(Orders!R:R)` |

## 🛠️ Skills Demonstrated

- Data cleaning: duplicates, blank rows, inconsistent formats/fonts
- PivotTables & PivotCharts for multi-dimensional business analysis
- Lookup formulas: VLOOKUP, XLOOKUP, INDEX-MATCH
- Calculated fields: IF, nested IF, IFS, SUMIFS, COUNTIFS
- Dashboard design: KPI cards, slicers, conditional formatting
- Working with the Excel Binary Workbook (.xlsb) format for larger files

## 📎 Project Structure

| File | Checkpoint |
|---|---|
| Checkpoint-1.xlsx | 1 — Data cleaning |
| Checkpoint-2.xlsx | 2 — Pivot tables |
| Checkpoint-3.xlsb | 3 — Lookup formulas |
| Checkpoint-4.xlsb | 4 — Calculated fields |
| Checkpoint-5.xlsb | 5 — Dashboard & conditional formatting |
| Screenshot.png | Dashboard preview |

## 👤 Author

Zahra Aliyeva
