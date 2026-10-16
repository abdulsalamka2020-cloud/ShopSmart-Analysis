![ShopSmart Retail Analytics: inside a ShopSmart store](/cover.jpg)

# ShopSmart Retail Sales Analysis with Microsoft Excel

**AnalystLab Africa Experience Lab | Week 1 | Data Analytics Track | Mission 01**
**Client:** ShopSmart Retail Ltd | **Role:** Junior Data Analyst | **Prepared by:** Abdulsalam | **Period analysed:** January to December 2025

---

## 1. Project Overview

ShopSmart Retail Ltd operates 25 stores across Nigeria and processes thousands of orders every month. Management has noticed fluctuating sales, inventory shortages and rising product returns. This project cleans the company's retail data, analyses it in Microsoft Excel, presents the results in an executive dashboard and turns them into business recommendations.

**Headline results**

| Metric | Value |
|---|---|
| Total revenue (all order statuses) | ₦6,818,849,032 |
| Total orders | 50,000 |
| Total customers | 5,000 |
| Average order value | ₦136,377 |
| Return rate | 4.88% |
| Inventory value (stock on hand x selling price) | ₦4,595,268,089 |

## 2. Business Problem and Objectives

1. Understand how the nine ShopSmart datasets relate to each other.
2. Clean and structure the data so it can be trusted.
3. Calculate the core sales KPIs with Excel formulas.
4. Combine the tables with lookup functions.
5. Summarise performance with Pivot Tables and Pivot Charts.
6. Present everything in an interactive Executive Dashboard.
7. Answer the ten business questions and give actionable recommendations.

## 3. Dataset

Nine CSV files plus a data dictionary were provided. The data dictionary documents five of the nine tables (Customers, Orders, Products, Stores, Returns); the keys of the other four were identified from their column headers.

| Sheet / Table | Represents | Rows | Primary key | Foreign keys |
|---|---|---|---|---|
| CUSTOMERS (`CTable`) | Customer profiles, location, membership segment | 5,000 | Customer_ID | none |
| PRODUCTS (`PTable`) | Product catalogue, category, brand, prices | 500 | Product_ID | none in this sheet |
| ORDERS (`OTable`) | Order header: date, status, channel, payment | 50,000 | Order_ID | Customer_ID, Product_ID, Store_ID |
| ORDER ITEMS (`OITable`) | Order line: quantity, price, discount, line total | 50,000 | Order_Item_ID | Order_ID, Product_ID |
| STORES (`STable`) | Store name, region, state, city, opening date | 25 | Store_ID | none |
| RETURNS (`RTable`) | Returned orders, reason, refund amount | 2,442 | Return_ID | Order_ID |
| EMPLOYEE (`ETable`) | Staff, role, store, salary | 125 | Employee_ID | Store_ID |
| SUPPLIERS (`SPTable`) | Supplier contact details | 489 | Supplier_ID | none |
| INVENTORY (`INTable`) | Stock on hand and reorder level per product | 500 | Inventory_ID | Product_ID, Store_ID |

**Relationships (one-to-many, from the first table to the second)**

```
CUSTOMERS ──< ORDERS >── STORES ──< EMPLOYEE
                │  │         └────< INVENTORY >── PRODUCTS
PRODUCTS ───────┘  ├──< ORDER ITEMS
                   └──< RETURNS
SUPPLIERS (supplier details, linked to products through the supplier name/brand)
```

## 4. Tools and Skills Used

Microsoft Excel: Excel Tables, SUM / AVERAGE / MAX / MIN / COUNTA / SUMPRODUCT, TEXT, CONCATENATE, IF, COUNTIF, VLOOKUP, Pivot Tables, Pivot Charts, Slicers, shapes linked to cells (KPI cards) and dashboard design.

## 5. Workbook Structure (`ShopSmart_Analysis.xlsx`)

| Sheet | Purpose |
|---|---|
| CUSTOMERS, PRODUCTS, ORDERS, STORES, RETURNS, EMPLOYEE, SUPPLIERS, ORDER ITEMS, INVENTORY | Cleaned data, each converted to an Excel Table |
| FORMULAS | Part 3 KPI formulas |
| LOOKUP | Part 4 lookup functions |
| PIVOT TABLE | Part 5 pivot tables (16) |
| CHARTS | Part 6 charts (6) |
| DASHBOARD | Part 7 executive dashboard |
| Sheet20, Sheet21 | Drill-through working sheets from the pivots |

## 6. Methodology: Step by Step

### Step 1: Data understanding
- Imported all nine CSV files into one workbook, one sheet per file, and opened each to inspect columns and data types.
- Read the data dictionary to understand each field, then identified the primary key of every table and the foreign keys that connect them (section 3).
- Mapped the brief's wording to real columns: "Customer Gender" is `Customers.Gender` and "Membership Level" is `Customers.Customer_Segment`.

### Step 2: Data cleaning
1. **Renamed worksheets** to match their contents (CUSTOMERS, PRODUCTS, ORDERS, and so on).
2. **Converted every dataset into an Excel Table** (`CTable`, `PTable`, `OTable`, `STable`, `RTable`, `ETable`, `SPTable`, `OITable`, `INTable`), so formulas use structured references and pivots refresh from the full table.
3. **Removed duplicate rows.** RETURNS had 2,500 records but only 2,442 unique orders, so 58 duplicate return records were removed. No other table contained duplicate rows or duplicate primary keys.
4. **Checked missing values** in all nine sheets. No blank cells remain in any table.
5. **Checked data types.** Dates (`Order_Date`, `Registration_Date`, `Return_Date`, `Opening_Date`, `Hire_Date`, `Last_Restock_Date`) are stored as true Excel dates, and all price, quantity, discount, cost and refund columns are numeric.
6. **Created helper columns** for analysis:

| Table | Column | Formula |
|---|---|---|
| CUSTOMERS | Full Name | `=CONCATENATE([@First_Name]," ",[@Last_Name])` |
| ORDERS | Order_Month | `=TEXT([@Order_Date],"mmm")` |
| ORDERS | Order_Year | `=TEXT([@Order_Date],"yyyy")` |
| ORDER ITEMS | Revenue | `=Quantity*Unit_Price` (gross) |
| ORDER ITEMS | Sales Amount | `=Line_Total*(1-Discount/100)` (net of discount) |
| ORDERS | Sales Amount | pulled from ORDER ITEMS |
| ORDERS | Category | `=VLOOKUP([@Product_ID],PTable[[Product_ID]:[Category]],3,FALSE)` |
| ORDERS | Is Returned | 1 if the order appears in RETURNS, otherwise 0 |
| INVENTORY | Inventory Value | `=Stock_On_Hand * Selling Price` |

7. **Corrected a Category lookup error.** The first version of the Category formula picked up the wrong product row and left most rows blank, which distorted the return rate by category. It was rewritten to use each order's own `Product_ID`, and all 50,000 rows now carry a valid category.

### Step 3: Excel formulas (sheet `FORMULAS`)

| # | KPI | Formula | Result |
|---|---|---|---|
| 1 | Total Revenue | `=SUM('ORDER ITEMS'!I:I)` | ₦6,818,849,032 |
| 2 | Total Quantity Sold | `=SUM('ORDER ITEMS'!D:D)` | 149,923 |
| 3 | Average Selling Price | `=AVERAGE(PRODUCTS!G:G)` | ₦36,244 |
| 4 | Average Order Value | `=Total Revenue / Number of Orders` | ₦136,377 |
| 5 | Highest Sales Amount | `=MAX('ORDER ITEMS'!I:I)` | ₦499,660 |
| 6 | Lowest Sales Amount | `=MIN('ORDER ITEMS'!I:I)` | ₦862 |
| 7 | Number of Orders | `=COUNTA(OITable[Order_ID])` | 50,000 |
| 8 | Number of Customers | `=COUNTA(CTable[Customer_ID])` | 5,000 |
| 9 | Number of Products | `=COUNTA(PTable[Product_ID])` | 500 |
| 10 | Number of Stores | `=COUNTA(STable[Store_ID])` | 25 |
| + | Inventory Value | `=SUMPRODUCT(INTable[Stock_On_Hand],PTable[Selling Price])` | ₦4,595,268,089 |
| + | Returned Orders | `=COUNTA(RTable[Return_ID])` | 2,442 |
| + | Return Rate | `=Returned Orders / Number of Orders` | 4.88% |

### Step 4: Lookup functions (sheet `LOOKUP`)
`VLOOKUP` was used to retrieve, from the relevant table, each of the seven required fields: **Customer Name** (from CUSTOMERS), **Store Name** (from STORES), **Product Name**, **Category** and **Brand** (from PRODUCTS), **Selling Price** (from ORDERS) and **Supplier Name** (from SUPPLIERS), all by matching on the ID column.

### Step 5: Pivot tables (sheet `PIVOT TABLE`)
Sixteen pivot tables were built from the Excel Tables:

| Required (Part 5) | Built from | Measure |
|---|---|---|
| Sales by Month | Order_Month | Sum of Sales Amount |
| Sales by Store | Store_Name | Sum of Sales Amount |
| Sales by Category | Category | Sum of Sales Amount, plus a count of orders |
| Sales by Brand | Brand | Sum of Sales Amount |
| Sales by State | State | Sum of Sales Amount |
| Sales by Customer Gender | Gender | Sum of Sales Amount, plus a customer count |
| Sales by Membership Level | Customer_Segment | Sum of Sales Amount |
| Top 10 / Bottom 10 Products | Product_Name | Value Filter (Top 10 / Bottom 10) on Sum of Sales Amount |
| Top 10 Customers | Customer Full Name | Value Filter (Top 10) on Sum of Sales Amount |
| Sales by Region (for Charts and Dashboard) | Region | Sum of Sales Amount |
| Return Rate by Category (for Charts) | Category | Average of Is Returned |
| Supplier and inventory summaries | Supplier name, Store Name | Sales, Inventory Value |

![Pivot tables from the PIVOT TABLE sheet](/pivot_tables.png)

**Reading the pivot tables:** each one is laid out as a labelled block with a green title bar, a filter button on the row field and a Grand Total row. The Sales by Month, Store, Category, State, Customer Gender, Membership Level and Region tables all reconcile to the same grand total of **₦6,818,849,032**, which confirms that no sales were lost or double counted in any grouping. The Top 10 and Bottom 10 Products tables show the best and weakest sellers (₦182.3M and ₦93.6M combined), and the Returned Rate table reports the average of the `Is Returned` flag for each category.

### Step 6: Charts (sheet `CHARTS`)
Six Pivot Charts were built on the `CHARTS` sheet. In each bar chart the highest bar is highlighted in a darker green so the leader stands out immediately.

![Six pivot charts from the CHARTS sheet](/charts.png)

| Chart | Type | What it shows | What it tells us |
|---|---|---|---|
| **Monthly Sales Trend** | Area | Total sales for each month, January to December 2025 | Sales are steady all year. February is the lowest month (₦509.5M) and December the highest (₦593.1M). The vertical axis starts at ₦460M, which makes the swings look bigger than they are: the whole range is about ₦83M on roughly ₦570M a month |
| **Revenue by Category** | Column | Total sales for each product category | Beauty leads with ₦1.59bn (23.3% of sales), followed by Groceries (₦1.42bn), Electronics (₦1.38bn), Fashion (₦1.27bn) and Home (₦1.16bn) |
| **Revenue by Store** | Bar | Total sales for each store name | Lagos, Ibadan and Abuja each sell about ₦1.1bn, while Enugu, Awka and Kano sell about ₦270M. The bars follow the number of outlets under each name (4 outlets for Lagos, Ibadan and Abuja; 3 for Asaba and Kaduna; 2 for Abeokuta and Port Harcourt; 1 each for Enugu, Awka and Kano), so individual outlets perform very similarly |
| **Top Products** | Column | The ten products with the highest sales | Home Item 400 leads at ₦19.76M. The top ten sit between about ₦17.4M and ₦19.8M and each is under 0.3% of total sales, so no single product carries the business. The axis starts at ₦16M, which exaggerates the gaps |
| **Sales by Region** | Doughnut | Share of total sales by region | South is largest at ₦2.18bn (32.0%), then East ₦1.91bn (28.0%), North ₦1.65bn (24.2%) and West ₦1.08bn (15.9%). The differences follow the number of stores in each region |
| **Return Rate by Category** | Pie | The share of orders returned within each category | Home (5.13%) and Beauty (5.08%) have the highest return rates and Groceries the lowest (4.58%), against a company average of 4.88%. Each slice is a category's own return rate, so the slices do not add up to 100% |

### Step 7: Executive dashboard (sheet `DASHBOARD`)
The dashboard brings the key numbers, charts and filters onto a single screen so a manager can read the state of the business at a glance.

![ShopSmart Retail Ltd Executive Overview dashboard](/executive_dashboard.png)

**Section 1: KPI cards (top row).** Six cards linked live to the FORMULAS sheet and the pivot tables, so they update when a slicer is used:

| Card | Value | Meaning |
|---|---|---|
| Total Revenue | ₦6,818,849,032 | Total sales after discount, all order statuses |
| Total Orders | 50,000 | Number of orders placed in 2025 |
| Total Customers | 5,000 | Registered customers |
| Avg Order Value | ₦136,377 | Revenue divided by number of orders |
| Return Rate | 4.88% | Returned orders as a share of all orders |
| Inventory Value | ₦4,595,268,089 | Stock on hand valued at selling price |

**Section 2: Charts (centre).**

| Chart | How to read it |
|---|---|
| **Monthly Sales Trend** | Shows seasonality. Sales dip in February and peak in December |
| **Revenue by Store** | Compares trading names. Lagos, Ibadan and Abuja are largest because each name covers four outlets |
| **Top Products** | Shows the ten best sellers. Home Item 400 is the leader (highlighted) |
| **Revenue by Category** | Shows where the money comes from. Beauty (highlighted) is the largest category |
| **Revenue by State** | Ranks states by sales. Lagos (₦1.11bn), Oyo (₦1.11bn) and Abuja (₦1.10bn) lead; Enugu, Anambra and Kano are lowest at about ₦270M each |

**Section 3: Interactive slicers (right-hand panel).** Click any button to filter the dashboard; click the clear-filter icon in a slicer's top corner to reset it.

| Slicer | Use it to |
|---|---|
| Order_Year | Choose the year (the data covers 2025 only) |
| Order_Month | Focus on one or more months |
| Category | Look at a single product category |
| Region | Compare East, North, South and West |
| Store_Name | Drill into a single store name |
| Customer_Segment | Compare Bronze, Silver and Gold members |

**Design choices:** one green colour scheme throughout, a darker green to highlight the top performer in each chart, icons on the KPI cards, rounded card layouts and clear chart titles. The sheet is exported to `Executive_Dashboard.pdf`.

### Step 8: Business questions

| # | Question | Answer |
|---|---|---|
| 1 | Product with the highest revenue | **Home Item 400**: ₦19.76M, followed by Groceries Item 115 (₦19.18M) and Groceries Item 331 (₦19.06M) |
| 2 | Best performing store | By store name, **ShopSmart Lagos** (₦1.11bn), but that name covers four outlets. By individual outlet, STR015 (Lagos) leads at ₦280.7M and STR002 (Kano) is lowest at ₦265.1M. All 25 outlets sit within about 6% of each other |
| 3 | Category that sells the most | **Beauty**: ₦1.59bn (23.3% of sales) and 11,809 orders, ahead of Groceries (₦1.42bn) |
| 4 | Month with the highest revenue | **December**: ₦593.1M (lowest: February, ₦509.5M) |
| 5 | Customers who spend the most | Ranked by Customer_ID: **Margaret Hawkins** (₦4.00M), Andrew Martin (₦3.85M), Joseph Collins (₦3.72M) |
| 6 | Products frequently returned | **Home Item 263** (11 returns), then Beauty Item 443, Home Item 5, Fashion Item 349, Beauty Item 262 and Beauty Item 220 (10 each) |
| 7 | Supplier supplying the most products | **No clear leader (a tie).** There are 489 suppliers for 500 products and each supplier maps to one product, so supplier revenue is used only as a tiebreaker in the discussion below |
| 8 | Stores with low inventory | Lowest total stock: **STR011 Kaduna (2,570 units)**, STR023 Abuja (3,139) and STR005 Enugu (3,349). 21 stock lines are at or below reorder level, mostly at Ibadan (5) and Kaduna (4) outlets |
| 9 | Best performing region | **South**: ₦2.18bn (32%); West is lowest at ₦1.08bn. The gap follows store count (South has 8 outlets, West 4), not outlet performance |
| 10 | Category management should invest in | **Beauty**: highest total sales (₦1.59bn) and highest completed-order revenue (₦523M), with a return rate (5.08%) close to the company average (4.88%) |

Other results: sales by gender are Male ₦3.49bn and Female ₦3.33bn; by membership level Gold ₦2.31bn, Silver ₦2.29bn and Bronze ₦2.22bn; return rate by category ranges from 4.58% (Groceries) to 5.13% (Home).

### Step 9: Business report (`Business_Report.pdf`)
**Three key findings**
1. Only about one third of reported revenue is from completed orders: ₦2.25bn of ₦6.82bn (33%). Cancelled orders are ₦2.30bn (34%) and Pending orders ₦2.28bn (33%).
2. Sales are spread evenly. Beauty leads (23.3%), monthly sales stay between ₦510M and ₦593M, and every outlet sells between ₦265M and ₦281M.
3. Returns are modest (4.88%, refunds of ₦61.8M, 0.9% of sales) and 75% are caused by operations: Late Delivery 26%, Wrong Item 25% and Damaged 24%.

**Three recommendations**
1. Fix order fulfilment first: clear the pending backlog and halve cancellations within 30 days, after breaking cancellations down by channel, payment method and store.
2. Cut avoidable returns with delivery-time standards, pick-and-pack checks and better packaging, aiming for a return rate below 4%.
3. Reorder the 21 low-stock lines, starting with Ibadan and Kaduna, and keep Beauty consistently in stock.

**Risk:** ₦4.60bn is tied up in inventory (about 67% of reported sales) while 21 stock lines are already at or below reorder level.
**Opportunity:** the 16,610 pending orders are worth ₦2.28bn, and every 10% converted adds about ₦228M.

## 7. Methodology Notes and Limitations

- **Revenue definition:** Sales Amount is the order-line total after discount, excluding shipping. Headline revenue includes all order statuses (Completed, Pending, Cancelled); the Completed-only figure is ₦2.25bn.
- **Line totals:** `Line_Total` in the source differs slightly from Quantity x Unit_Price on many rows (₦7.575bn vs ₦7.576bn in total), so revenue is based on the recorded `Line_Total`.
- **Stores:** there are 25 store IDs but only 10 store names, because names such as "ShopSmart Lagos" repeat across regions. Store-level pivots by name therefore combine several outlets.
- **Regions** are taken from the Stores table as given and do not always match the store's city or state, so regional conclusions should be read with that in mind.
- **Suppliers:** the data has no product-to-supplier link that identifies a leading supplier, so supplier revenue could not be attributed reliably.
- **Customers:** customer names repeat (180 duplicate full names), so customer rankings use `Customer_ID` rather than name.
- **Inventory value** uses selling price, not cost price, so it is a retail value.
- **Year filter:** the data covers 2025 only, so the Year slicer has one option.

## 8. Future Improvements (Business Strategy)

1. Make order completion rate a headline KPI alongside revenue, so management tracks revenue that is actually delivered.
2. Review the supplier base: with one supplier per product, ShopSmart has no supplier concentration risk but also no purchasing leverage, so consolidating suppliers could lower costs.
3. Add a profit view (cost price against selling price) to see which categories earn the most margin before shifting investment.
4. Track stock cover by store against sales so reorder levels reflect actual demand.
5. Capture delivery-time data to measure the link between late delivery and returns.

## 9. Repository Structure

```
Week1_Abdulsalam_<LastName>/
├── ShopSmart_Analysis.xlsx
├── Executive_Dashboard.pdf
├── Business_Report.pdf
├── README.md
└── images/
    ├── cover.jpg
    ├── pivot_tables.png
    ├── charts.png
    └── executive_dashboard.png
```

The `images/` folder holds the pictures shown in this README. Keep it in the same folder as `README.md` so the images display on GitHub.

## 10. Author

**Abdulsalam**, Junior Data Analyst intern, AnalystLab Africa Experience Lab (Week 1).
Tags: #AnalystLabAfrica
