# Superstore Sales Dashboard (Excel)

This is an Excel dashboard I built using Superstore sales data with around 10K order lines from 2023 to 2026.

I mainly wanted to understand where the profit is coming from, where it is being lost, and how discounts, products and regions affect overall performance.

![Dashboard](Screenshot.png)

## Main findings

Sales are growing, but profit is growing more slowly.

* **2026 sales increased by 21.4%** compared with 2025, while **profit increased by 16.0%**. As a result, the profit margin went down from 13.5% to 12.9%.
* **Discounts above 20% are a major problem.** These 1,415 order lines generated $365K in sales but resulted in a **$136K loss**. Orders with no discount have a 29.6% margin, while orders with discounts of 40% or more lose 77 cents for every dollar of sales.
* **Furniture has strong sales but very little profit.** It generates $755K in sales, which is similar in scale to the other categories, but only $20K in profit. Its margin is 2.6%, compared with around 17% for Office Supplies and Technology.
* **Tables and Bookcases are the main loss makers.** Tables lost $17.8K and Bookcases lost $3.6K. They also have the highest average discounts in the Furniture category, at 25.8% and 21.5%.
* **Some states perform much worse than others.** Texas had a $25.7K loss, followed by Ohio with $17.0K and Pennsylvania with $15.6K. Looking at regions, Central has the lowest margin at 7.9%, while West has the highest at 15.0%.

## Key numbers

| KPI           | Value                      |
| ------------- | -------------------------- |
| Total Sales   | $2.33M                     |
| Total Profit  | $292.3K                    |
| Profit Margin | 12.6%                      |
| Total Orders  | 5,111 (distinct Order IDs) |

The growth figures on the KPI cards compare 2026 with 2025. The main KPI values change when the slicers are used.

## Dashboard

The dashboard includes:

| Visual                                   | What it shows                                                                                   |
| ---------------------------------------- | ----------------------------------------------------------------------------------------------- |
| KPI cards                                | Sales, profit, margin and orders, with the 2026 vs 2025 change                                  |
| Sales and Profit by Month                | How sales and profit change throughout the year, with sales reaching their highest levels in Q4 |
| Discounts Above 20% Turn Profit Negative | Profit across different discount bands                                                          |
| Bottom 5 Sub-Categories by Profit        | The sub-categories with the lowest profit, including Tables, Bookcases and Supplies             |
| Sales vs Profit by Category              | Sales, profit and margin for each category                                                      |
| Profit by State                          | Profit and loss across different states                                                         |
| Slicers                                  | Segment, Region and Year                                                                        |

## Recommendations based on the analysis

1. **Keep discounts at or below 20%.** Discounts above this level are not profitable on average and appear to be one of the main areas affecting profit.
2. **Review Tables and Bookcases.** Their losses are large enough to significantly affect the overall Furniture result, so their pricing and discount levels need closer attention.
3. **Investigate Texas, Ohio and Pennsylvania.** It would be useful to check whether the losses in these states are mainly related to discounts, product mix or shipping costs.
4. **Look at profit together with sales.** Technology makes $147K profit from $840K in sales, while Furniture makes only $20K from $755K in sales.

## How I built it

I started by cleaning the raw data, removing 10 duplicate rows and one blank row. I also converted Order Date from text into a proper date, fixed number formats, and standardized the headers and fonts.

After that, I created PivotTables to look at sales and profit by sub-category, category, region, ship mode, year and discount band.

For the lookup part, I used **INDEX-MATCH and XLOOKUP** for Regional Manager and cross-checked the results. I also used lookups for Supplier, Unit Cost and Target Margin, as well as a VLOOKUP flag for returned orders.

I created several calculated fields, including **Profit Status (IF), Discount Level (nested IF), Margin Category (IFS)** and **Discount Band**. I also used **SUMIFS and COUNTIFS** for some of the regional summaries.

The final dashboard was built with PivotCharts, KPI cards linked to cells, three connected slicers and conditional formatting to make profit and loss easier to see.

## A few things I learned from the project

One issue I came across was with the **Total Orders** calculation. A helper column that counted only single-line orders gave me 2,593 instead of the correct 5,111. Using the distinct Order IDs fixed the problem. It was a good reminder to always check an important KPI against the raw data.

I also noticed that when month and quarter data from all years are shown together, it tells you more about **seasonality** than about yearly growth. Because of that, I used year-over-year figures for the growth badges instead of presenting the monthly view as a trend.

Another small lesson was that chart titles are more useful when they explain what the viewer should notice, rather than simply repeating the name of the metric.

**Tools:** Excel (PivotTables, PivotCharts, slicers, XLOOKUP, INDEX-MATCH, IFS
