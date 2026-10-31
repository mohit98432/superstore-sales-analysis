# Superstore Sales Analysis Dashboard

## Objective
Explore sales, profitability, seasonality, customers and discount performance in an Excel portfolio project for entry-level data analyst roles.

## Dataset
Sample Superstore: 9,994 order lines, 5,009 distinct orders, covering 2014–2017. Currency: USD. This is a sample dataset, not a current business report.
Public source: https://github.com/eyowhite/Data-Analysis/blob/main/Sample%20-%20Superstore.csv

## Steps and tools
1. Downloaded the CSV and retained original values on Raw Data.
2. Used Microsoft Excel automation to remove exact duplicates across all source fields, check blanks, convert numeric fields and dates, and check shipping chronology. Cleaning Log records results. Repeated Order IDs are valid because an order can contain multiple products.
3. Added Excel formulas for Month, Year, line Profit Margin and Order Weight.
4. Created native PivotTables on separate sheets for region, category, sub-category, monthly sales by year, top products, top customers and discount. A supporting pivot calculates KPIs.
5. Created five dashboard charts and Region/Year slicers connected to all eight pivots.
6. Added findings and recommendations. Used the bundled spreadsheet library for initial workbook structure and Excel for native features.

## Results
- Sales: $2,297,200.86; profit: $286,397.02; orders: 5,009; profit margin: 12.47%.
- West leads regional sales and profit.
- Technology contributes the most category profit.
- Furniture margin is approximately 2.5%, despite substantial sales.
- Tables generate approximately $17,725 in losses.
- Sales increased approximately 51.4% between 2014 and 2017.

## How to use
Open the workbook in desktop Excel. Select Region and Year slicers; use the clear-filter icon to restore all data. Review the underlying pivots to explain charts. Summary findings describe the full dataset and do not change with slicers. Refresh All after changing source data, and extend helper formulas when adding rows.

## Definitions and limitations
Overall margin = sum of profit / sum of sales. Order Weight = 1 / number of lines in an order. Summing this counts distinct orders under Region and Year filters because each order belongs to one region and year. Do not apply this method unchanged to category/product filters. Top 10 lists group by displayed names; investigate IDs before operational customer/product decisions. Discount association is descriptive and does not establish causality. No real-world revenue improvement is claimed.

## Resume bullets
- Built an Excel sales analysis dashboard using 9,994 order lines and 5,009 distinct orders, with four KPIs, five charts and connected Region/Year slicers.
- Validated source data, converted date and numeric fields, added calculated columns and documented cleaning decisions in an auditable cleaning log.
- Created eight native PivotTables supporting regional, product, customer, monthly and discount analysis, including top-10 sales rankings.
- Identified low furniture margins and loss-making tables, translating findings into three business recommendations.

Use these bullets after you have practiced navigating the workbook and can explain the calculations. The build was assisted by automation.

## Interview practice: 10 questions and answers
1. **What was the objective?** Identify where sales and profit come from, how sales change over time and where profitability needs investigation.
2. **What does one row represent?** One order line. An order can have several lines, so counting rows overstates orders.
3. **How did you handle duplicates?** Checked exact duplicates across all 21 original fields in Excel. Repeated order IDs were retained as legitimate multi-product orders.
4. **How did you handle dates?** Interpreted the source as US month/day/year, converted to Excel date values and applied a readable date format. Formatting text alone is insufficient.
5. **How did you calculate margin?** Total profit divided by total sales. An unweighted average of row margins would overrepresent small transactions.
6. **Why use PivotTables?** They aggregate detailed records into reusable summaries and support connected slicer filtering without rewriting formulas.
7. **How does the order KPI work?** Each line receives 1 divided by its order's line count. These weights sum to one per order. This is valid for the supplied Region/Year filters, not arbitrary product filters.
8. **What was your strongest insight?** Furniture generated substantial revenue but only about 2.5% margin. Tables lost money, suggesting a targeted pricing and cost review.
9. **Do discounts cause losses?** This analysis cannot establish causality. Product mix, regional differences and costs may explain part of the relationship. I would compare similar products and test policy changes.
10. **How did you validate the dashboard, and what would you improve?** Reconciled full-data KPIs, checked native objects and tested West filtering against its independent sales total. Next I would add Power Query refresh, ID-based product/customer dimensions and a data-model distinct-order measure for broader filters.

## Learning and review update
Read Excel_Project_Learning_Guide.md for worked formulas, pivot recipes, filter checkpoints, exercises and a two-minute interview explanation. The review repaired clipped KPI cards and 65 incorrectly decoded source strings, reconciled all raw source fields, checked cached formula errors and tested combined West/2017 filtering in desktop Excel. See Cleaning Log for the audit trail.
