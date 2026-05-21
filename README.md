# Data Cleaning & Reporting Automation

An automated data cleaning and reporting pipeline built using Python for Walmart sales data analysis.

---

## Project Overview

This project automates the complete data preprocessing and reporting workflow using Python.

The system processes Walmart sales data from 45 stores (2010–2012) and automatically performs:

- Data Cleaning
- Missing Value Handling
- Duplicate Removal
- Outlier Detection
- Feature Engineering
- Data Normalization
- Dashboard Visualization
- Excel Reporting Automation

---

## Features

- Automated Data Cleaning Pipeline
- Missing Value Imputation
- Duplicate Detection & Removal
- Outlier Detection using IQR
- Feature Engineering
- Data Normalization
- KPI Dashboard Visualization
- Seasonal Sales Analysis
- Store Performance Analysis
- Correlation Analysis
- Automated Excel Report Generation

---

## Technologies Used

- Python
- Pandas
- NumPy
- Matplotlib
- OpenPyXL

---

## Dataset Information

Dataset includes:

- Store
- Date
- Weekly Sales
- Holiday Flag
- Temperature
- Fuel Price
- CPI
- Unemployment

---

## Data Cleaning Pipeline

The pipeline performs:

1. Date Parsing
2. Duplicate Removal
3. Missing Value Handling
4. Data Type Correction
5. Outlier Detection
6. Feature Engineering
7. Feature Normalization

---

## Outlier Detection Formula

Uses IQR Method:

:contentReference[oaicite:0]{index=0}

Lower Bound:

:contentReference[oaicite:1]{index=1}

Upper Bound:

:contentReference[oaicite:2]{index=2}

---

## Dashboard Components

The generated dashboard contains:

- Cleaning Pipeline Summary
- KPI Cards
- Monthly Sales Trends
- Holiday Impact Analysis
- Top 10 Stores Revenue
- Feature Correlation Analysis

---

## Project Structure

```text
Data-Cleaning-and-Reporting-Automation/
│
├── automation_pipeline.py
├── sales_data.csv
├── data_cleaning_dashboard.png
├── walmart_cleaned.csv
├── walmart_data_cleaning_report.xlsx
└── README.md
```

---

## Installation

Install required libraries:

```bash
pip install pandas numpy matplotlib openpyxl
```

---

## How to Run

```bash
python automation_pipeline.py
```

---

## Generated Outputs

The project automatically generates:

### 1. Dashboard Visualization

```text
data_cleaning_dashboard.png
```

Contains:
- KPI Cards
- Store Analysis
- Monthly Trends
- Correlation Charts

---

### 2. Cleaned Dataset

```text
walmart_cleaned.csv
```

Contains cleaned and transformed data.

---

### 3. Automated Excel Report

```text
walmart_data_cleaning_report.xlsx
```

Includes:
- Cleaned Data
- Yearly Summary
- Monthly Summary
- Store Performance
- Cleaning Log
- Automated Charts

---

## Business Insights

- Analyze seasonal sales patterns
- Identify top-performing stores
- Understand holiday sales impact
- Detect sales anomalies
- Monitor economic factor influence on sales

---

## Future Improvements

- Streamlit Interactive Dashboard
- Power BI Integration
- Real-Time Data Pipeline
- Predictive Analytics
- Automated PDF Reports
- Machine Learning Forecasting

---

## GitHub Repository Name

```text
Data-Cleaning-and-Reporting-Automation
```

---

## Author

Your Name

---

## License

MIT License
