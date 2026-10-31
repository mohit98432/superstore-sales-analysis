# Superstore Sales Analysis — learn and present the project

## Start here
Open Superstore_Sales_Analysis.xlsx in desktop Excel. Read this guide beside it. Work through one section at a time; the aim is to explain why each step is needed, not memorise button clicks.

The question: **Where are sales strong, where is profit weak, and what should a retail manager investigate?**

This is historical sample data for 2014–2017, in US dollars. It is a practice project, not evidence of an actual company's current performance. Your Excel regional settings may display $22,97,201 instead of $2,297,201. Both represent the same number.

## 1. Understand the rows before calculating anything
Open Raw Data. Each row is one product line within an order. The first two records share Order ID CA-2016-152156 but contain different products. Both must stay.

Three different counts answer three different questions:
- 9,994 rows = order lines.
- 5,009 distinct Order IDs = orders.
- Sum of Quantity = units sold, which is another measure entirely.

Important fields:
| Field | Meaning | Treatment |
|---|---|---|
| Row ID | Source row identifier | Not a business measure |
| Order ID | Identifier shared by all lines of an order | Text; repeated values can be valid |
| Order Date / Ship Date | Purchase and shipment dates | Real Excel dates |
| Customer ID / Product ID | Entity identifiers | Text; retain for investigation |
| Customer Name / Product Name | Human-readable labels | Used in the current top-10 lists |
| Region / Category / Sub-Category | Grouping dimensions | Use in pivot Rows or filters |
| Sales | Recorded line sales, USD | Sum; do not multiply by Quantity again |
| Profit | Recorded line profit, USD | Sum; losses are meaningful |
| Discount | Recorded discount rate | 0.20 means 20% |
| Postal Code | Geographic identifier | Text, not a quantity |

The data already supplies Sales and Profit. We are analysing those measures, not reconstructing accounting profit from missing cost records.

## 2. Cleaning: what actually changed
Raw Data retains the imported source representation; Clean Data contains typed fields and helper formulas. The Cleaning Log documents the original build.

An independent source audit on 31 October 2026 found:
| Check | Result | Decision |
|---|---:|---|
| Source rows | 9,994 | Retain |
| Empty source cells | 0 | No imputation needed |
| Exact duplicate rows across all 21 fields | 0 | No rows removed |
| Ship Date earlier than Order Date | 0 | No chronology corrections |
| Zero Sales | 0 | No zero-denominator rows |
| Negative Profit lines | 1,871 | Keep; these are loss-making sales |
| Orders spanning multiple Region/Year combinations | 0 | Order-weight method valid for these two slicers |

Do not say you removed duplicates if the check found none. Say: "I checked for exact duplicates and found none. I preserved repeated order IDs because the data is at line level."

The numeric fields were converted to numbers. Source dates use month/day/year: 11/08/2016 means 8 November 2016. Display formatting alone does not turn text into a real date. The source CSV uses Windows-1252 text encoding. This review repaired 65 unique strings damaged by the earlier import, including their PivotTable cache entries. All 209,874 raw field values were then compared with the source and matched. Specify encoding explicitly when reproducing the import.

To reproduce cleaning manually on a practice copy:
1. Import the CSV with Data > From Text/CSV, inspect encoding and delimiter, and retain an untouched raw copy.
2. Check the row count and define what constitutes a duplicate before using Data > Remove Duplicates.
3. Check blank cells. Do not fill missing values without a business rule.
4. Convert dates using the source locale, then format them as dates.
5. Confirm Sales, Quantity, Discount and Profit are numbers. Keep identifiers as text.
6. Add formulas below, then record checks and transformations in a log.

## 3. Read the formulas
Open Clean Data and inspect the formula bar in columns V:Y.

### Month
`=MONTH(C2)`

C2 is Order Date. MONTH returns 1 through 12. Numeric months keep the trend in calendar order. Alphabetical month names can produce April before February.

### Year
`=YEAR(C2)`

Returns the year used by the Year slicer and monthly pivot.

### Line profit margin
`=IF(R2=0,0,U2/R2)`

R2 is Sales and U2 is Profit. The first row has sales 261.96 and profit 41.9136, so its margin is 16%. IF avoids division by zero; the present dataset has no zero sales. In a future dataset, consider returning a clearly labelled not-applicable result instead of treating a zero-sales margin as a genuine 0%.

### Order Weight
`=1/COUNTIF($B$2:$B$9995,B2)`

COUNTIF counts how many lines belong to the current order. If an order has two lines, each gets 0.5; together they sum to one order. The dollar signs keep the counting range fixed when copying the formula.

This is a workaround for counting orders in a conventional pivot. It works with this dashboard's Region and Year slicers because every order belongs to one Region and Year. It does **not** reliably count distinct orders after arbitrary category/product filtering. For that extension, use a Data Model distinct count of Order ID. When adding rows, extend the COUNTIF range as well as the table and formulas.

### Dashboard values
The dashboard uses GETPIVOTDATA to retrieve the filtered totals from KPI Support. For example:

`=GETPIVOTDATA("Total Sales",'KPI Support'!$A$4)`

This asks the supporting pivot for its Total Sales measure. It is more explicit than relying on a total always appearing in a particular row.

The dashboard margin is `=IF(B6=0,0,F6/B6)`: filtered total profit divided by filtered total sales.

Why not average line margins? Suppose one line sells $100 with $50 profit and another sells $900 with $90 profit. Average line margin is (50% + 10%)/2 = 30%. Actual combined margin is $140/$1,000 = 14%. The bigger sale must carry more weight.

## 4. Understand and rebuild each PivotTable
On a practice copy, click inside Clean Data, choose Insert > PivotTable, and use the SuperstoreData table. Add the following fields. Numeric Values must say **Sum**, not Count.

| Sheet | Rows | Columns | Values | Business question |
|---|---|---|---|---|
| Region | Region | Values | Sum of Sales, Sum of Profit | Which regions contribute most? |
| Category | Category | Values | Sum of Sales, Sum of Profit | Which broad product groups perform well? |
| Sub-category | Sub-Category | Values | Sum of Sales, Sum of Profit | Which detailed groups lose money? |
| Monthly Trend | Month | Year | Sum of Sales | How do months compare across years? |
| Top Products | Product Name | Values | Sum of Sales, Sum of Profit | Which names generate most sales? |
| Top Customers | Customer Name | Values | Sum of Sales, Sum of Profit | Which customers generate most sales? |
| Discount | Discount | Values | Sum of Sales, Sum of Profit | How does aggregate profit vary by discount? |
| KPI Support | Region | Values | Sum of Sales, Sum of Profit, Sum of Order Weight | What totals should cards display? |

For each top-10 pivot, sort descending by Total Sales and apply Value Filters > Top 10 by Total Sales. The current lists group by displayed names; IDs should be examined before making customer-specific or product-specific operational decisions.

The Month pivot contains numbers 1–12. Each line in its chart is a year. Comparing January with January is different from comparing December with January: seasonal changes can obscure the longer-term pattern.

The discount pivot shows totals. A larger total may reflect more transactions, so do not interpret the chart as a controlled experiment. Add weighted margin and transaction counts in a future extension to separate scale from profitability.

## 5. How the dashboard works
The five charts cover Region, Category, Monthly Trend, Sub-category and Discount. Top products and customers remain available on their own analysis sheets.

The two slicers filter connected PivotTables. Their shared source cache lets one selection update several views. The cards read KPI Support, so they update too. Summary text is deliberately a full-dataset narrative and does not recalculate with slicers.

Use the funnel-with-cross icon on **both** slicers to restore all data. Clearing Region alone leaves Year filtering active.

Confirmed checkpoints:
| Selection | Sales | Profit | Orders | Margin |
|---|---:|---:|---:|---:|
| All regions, all years | $2,297,200.86 | $286,397.02 | 5,009 | 12.47% |
| West, all years | $725,457.82 | $108,418.45 | 1,611 | about 14.94% |
| West, 2017 | $250,128.37 | $43,808.96 | 538 | about 17.51% |

Cards show rounded dollars and one-decimal percentages. Underlying totals retain more precision. Region and combined Region/Year selections were tested in desktop Excel against independent calculations.

## 6. Findings and recommendations
1. West leads regional sales and profit. It contributes about $725k sales and $108k profit.
2. Technology has the highest category sales ($836,154) and profit ($145,455).
3. Furniture contributes $741,999.80 sales but only $18,451.27 profit: approximately 2.49% margin. High sales alone do not establish a healthy business.
4. Tables lose about $17,725. Investigate pricing, discounts and costs at product level before deciding whether to discontinue a category.
5. Sales rose from $484,247 in 2014 to $733,215 in 2017, approximately 51.4%. Growth was not uninterrupted: 2015 sales were $470,533, below 2014.

Recommended actions:
- Prioritise loss-making tables for a pricing, discount and fulfilment-cost review.
- Test targeted Technology promotions while monitoring profit and margin alongside revenue.
- Compare discount outcomes within similar products and regions before changing discount rules. A small controlled test is stronger evidence than aggregate correlations.

These are proposed actions. Do not claim the project increased revenue or profit: no recommendations were implemented or measured.

## 7. Practice exercises with answers
1. Explain why 9,994 is not the number of orders. **Answer:** multiple product lines can belong to one Order ID.
2. Find the first line's margin. **Answer:** 41.9136 / 261.96 = 16%.
3. Set West and 2017. What happens? **Answer:** all connected pivots and the KPI cards filter to that intersection; expected values are in the checkpoint table.
4. Clear only Region. Is the whole dataset restored? **Answer:** no, the 2017 filter remains. Clear Year too.
5. Why keep negative profits? **Answer:** removing them would bias the analysis and hide losses.
6. Which category illustrates the difference between revenue and profitability? **Answer:** Furniture, with substantial sales but roughly 2.49% margin.
7. Did sales increase every year? **Answer:** no. They fell in 2015 before rising in 2016 and 2017.
8. Why does the order-weight formula need care when adding rows? **Answer:** its counting range is bounded. New order lines require the range and helper formulas to expand.
9. Can this analysis prove discounts caused losses? **Answer:** no. Product mix, geography and other factors are not controlled.
10. Make a new Region pivot on a separate practice sheet without following instructions. **Answer check:** all-region sales and profit should reconcile to the full-data checkpoint.

## 8. Two-minute interview explanation
"I analysed a retail sample dataset containing 9,994 order lines across 5,009 orders. I first established the row-level meaning, preserved the source and checked duplicates, missing values, date types and chronology. I added Month, Year, profit-margin and order-count helpers, then built eight PivotTables and a dashboard with five charts and connected Region and Year slicers.

I used total profit divided by total sales for overall margin rather than averaging line margins. I also distinguished order count from line count. The full dataset generated approximately $2.30 million in sales and $286,397 in profit, a 12.47% margin.

Furniture stood out because it had substantial revenue but only about 2.5% margin, while Tables lost money. I recommended a targeted pricing and cost review. I validated dashboard totals and tested combined filters. The work was assisted by automation; I can explain the formulas, reproduce the pivots and discuss limitations, including the order-weight method and the fact that discount associations do not prove causation."

## 9. What to study next
Use the 10 interview questions and four resume bullets in README.md. First practise rebuilding one pivot yourself, then reconstruct the margin and order-count logic. Only list skills you can demonstrate. A useful next version would add Power Query refresh, explicit source-encoding control, ID-based dimensions and a Data Model distinct-order measure.

## Review scope
This pass repaired the dashboard's clipped KPI values, independently audited source counts and main figures, and tested Region and combined Region/Year slicers in Excel. The source audit used Windows-1252 decoding and reconciled every raw source field after correcting text encoding. The workbook reopened successfully in desktop Excel. No cached formula error cells were found across all worksheets. The dashboard, Summary and Cleaning Log were visually inspected; a full visual inspection of every supporting worksheet was not performed.

