# Universal Excel Data Cleaner

A reusable Excel VBA automation for first-pass data cleaning and data quality reporting.

## Overview

Universal Excel Data Cleaner is a VBA-based automation designed to process different Excel datasets with a single click.

The macro is stored in Excel's `PERSONAL.XLSB` and works on the currently active workbook. It creates a cleaned output sheet and a data quality report without requiring the VBA code to be copied into every new dataset.

## Features

- Clean extra spaces and common text inconsistencies
- Identify missing values
- Detect duplicate records
- Validate email fields
- Normalize and check phone/contact fields
- Detect and format date-related fields
- Clean common financial fields such as salary, revenue, amount, price, and cost
- Generate data quality flags
- Create a `Clean_Data` output sheet
- Create a `Data_Quality_Report`
- Run the workflow using a single Quick Access Toolbar button

## How It Works

```text
Open Excel Dataset
       ↓
Activate Data Sheet
       ↓
Click "Clean Data"
       ↓
Universal_Data_Cleaner runs
       ↓
Clean_Data
       +
Data_Quality_Report
