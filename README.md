# FUTURE_DS_01 – Sales Data Analysis

## 📌 Task Overview
This project analyzes retail sales data to identify revenue trends, top-selling products, and high-revenue generating regions and categories. The goal is to simulate a real-world business analytics task (Future Interns – Data Science & Analytics, Task 1).

## 🧰 Tools Used
* Python (Pandas, NumPy, openpyxl)
* Microsoft Excel (formula-driven summary tables and charts)
* HTML dashboard, Power BI style (Chart.js)

## 📁 Dataset
Simulated retail sales dataset (self-created, as allowed by the task): 6,180 raw order lines, January 2023 – December 2025, with 4 regions, 3 categories and 14 products. It is built to include realistic data problems (duplicates, inconsistent text, missing values, mixed date formats).

## 🧹 Data Cleaning Steps
* Identified null values (Region) and invalid values (Quantity ≤ 0)
* Removed fully duplicated rows (180)
* Standardized Region text (trimmed spaces, fixed upper/lower case)
* Removed rows with missing Region (90) and invalid Quantity (38)
* Converted two date formats (YYYY-MM-DD and DD/MM/YYYY) into one
* Created Revenue = Quantity × UnitPrice × (1 − Discount), Profit = Revenue − Cost, plus Month and Year

Result: 5,872 clean rows (see `Cleaning_Log` sheet).

## 📊 Key Analysis
* Monthly revenue trends and year-over-year growth
* Top products by revenue
* Revenue and profit by category and region
* Profit margin by discount level

## 💡 Key Insights
* Total revenue 2.32M, profit 466K, margin 20.1%; revenue grew 16.2% (2024) and 11.1% (2025)
* Technology = 48.7% of revenue; Office Supplies has the best margin (43.3%); Furniture only 1.9%
* West is the largest region (32.2%); South has the lowest margin (17.4%)
* Q4 (Oct–Dec) brings 33.8% of yearly revenue
* 30% discounts are loss-making (−8.6% margin)

## ✅ Recommendations
Cap discounts at 10%, plan stock for September–December, bundle high-margin items with big-ticket products, and add discount-approval rules in the South region.

## 📈 Dashboard Preview
*(Add your screenshot here: `![Dashboard](dashboard.png)`)*

## 📂 Files
* `Sales_Performance_Dashboard.xlsx` – Excel dashboard, summary tables, insights
* `Sales_Dashboard_PowerBI_Style.html` – interactive one-page dashboard
* `sales_analysis.py`, `build_dashboard.py` – full pipeline
* `data/raw_sales.csv`, `data/clean_sales.csv`

## 📝 Notes
The dataset is simulated, so the numbers illustrate the analysis method rather than a real company. The workflow can be rerun on the Kaggle Online Retail or Superstore dataset by replacing the CSV and updating column names.
 ----end of the topic