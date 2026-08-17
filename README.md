# Excel Minor Project 5: Sales Operations Analytics 🛒📊

Welcome to the **Sales Operations Analytics** project repository! This project evaluates real-world retail enterprise data across 6 Indian cities using **Microsoft Excel** to answer critical business questions, diagnose data quality issues, and calculate key performance indicators (KPIs) for executive review.

---

## 📌 Business & Executive Summary

* **Total 2024 Revenue:** **₹26,572,948.20** across 2,000 order transactions.
* **Top Performing Region:** **South Region** (**₹7.14M** total revenue), followed closely by **North** (**₹6.98M**).
* **Top Product Category:** **Beauty** (**₹5.61M** total sales, accounting for ~21.1% of total revenue).
* **Data Hygiene Audit:** Identified **359 duplicate orders**, **30 broken foreign key references**, **174 orders for discontinued products**, and **58 promo-abuse warning cases**.

---

## 🔍 Core Analysis & Questions Solved

### Section 1: Lookups & Data Retrieval (25 Marks)
* **Q1.1 Customer Identification:** Mapped `order_id` `ORD-00234` to `CUST-0173` using `VLOOKUP`.
* **Q1.2 Chain Lookup:** Looked up product ID and retrieved name (`Dumbbells 5kg`) for `ORD-00456`.
* **Q1.3 Left-Lookup (INDEX-MATCH):** Retrieved sales rep name (`Nisha Menon`) for `ORD-00987`.
* **Q1.4 2D Lookup:** Retrieved target value (**₹321,255**) for rep `R012` in `Mar-24`.
* **Q1.5 Multi-hop Lookup:** Identified Category (`Electronics`) for `ORD-01500`.
* **Q1.6 Approximate Lookup:** Categorized discount slab for `ORD-00789` (`Bronze`).
* **Q1.7 Dynamic String Formatting:** Concatenated string `Sanjay Kapoor (South)` for `ORD-01234`.

---

### Section 2: Data Cleaning & Diagnostics (20 Marks)
* **Duplicate Rows:** Identified **359 duplicate order records**.
* **Whitespace Auditing:** Found **15 customer names** containing leading/trailing spaces.
* **Case Standardization:** Standardized city cases down to **6 distinct cities**.
* **Referential Integrity Check:** Flagged **30 orders** with invalid `customer_id` references.
* **Discontinued SKU Audit:** Detected **174 orders** associated with inactive products (`is_active = FALSE`).
* **Promo Abuse Diagnostic:** Flagged **58 orders** that were `Cancelled` with a `discount_pct > 20%`.

---

### Section 3: Conditional Aggregation & Business Metrics (25 Marks)
* **Regional Q3 Revenue:** **₹1,217,239.01** generated in Mumbai during Q3 2024.
* **High-Discount Completed Orders:** **723 orders** had `discount_pct > 15%` & `payment_status = Completed`.
* **Category/Segment Average:** **₹13,981.40** average order value for Corporate Electronics purchases.
* **Top 3 Sales Rep Revenue:** Combined revenue of **₹3,342,678.55** driven by the top 3 reps.
* **Target vs Actual Performance:** Evaluated rep `R015` performance tier for Nov-24 as **Below** target (<80%).
* **Age Segment Cohort:** **₹982,568.40** total revenue from Corporate customers aged 30–40.

---

### Section 4: Pivot Tables & Multi-Dimensional Analysis (20 Marks)
* **Regional & Category Matrix:** South region Electronics sales reached **₹1,314,107.39**.
* **Peak Monthly Rep Revenue:** Highest individual rep month contribution was **₹216,421.52**.
* **Category Contribution:** Apparel accounts for **14.02%** of North region total sales.
* **Quarterly Distribution:** Q4 2024 total revenue reached **₹6,459,442.98**.
* **Calculated Profit Margin:** Average calculated profit margin for Delhi shipments was **82.01%**.

---

### Section 5: Advanced Analytics & Formulas (10 Marks)
* **Rank #7 Sales Rep:** `Amit Singh`.
* **YTD Revenue (Jan-Aug 2024):** Cumulative revenue reached **₹18,050,689.60**.
* **West Region #2 Product:** `Smart Watch` ranked 2nd by total revenue in the West.
* **Top 5 Reps Concatenation:** `Neha Desai | Shalini Patel | Deepak Shah | Nisha Menon | Arun Iyer`.
* **Peak Order Value (Grocery/Bengaluru):** Maximum single order reached **₹22,134.00**.

---

## 🛠️ Tooling & Functions Used

* **Excel Functions:** `VLOOKUP`, `INDEX-MATCH`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, `LARGE`, `RANK.EQ`, `TEXTJOIN`, `TRIM`, `LEN`, `PROPER`
* **Features:** PivotTables, Calculated Fields, Grouping, Data Validation

---

## 🎓 Acknowledgments
Special thanks to **The Unlox Academy** for providing the dataset and assessment scenario!
