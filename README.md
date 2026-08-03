# Excel VBA: Data Cleaner & Pivot Table Generator

This repository contains a Visual Basic (VBA) module that automates the data preparation and analysis pipeline directly inside Microsoft Excel. It cleans raw datasets by removing unwanted rows, adds time-intelligence by tagging quarters and years, and dynamically generates interactive Pivot Tables for instant reporting.

- **Dynamic Data Cleaning**: Scans your dataset and automatically deletes unwanted or redundant rows based on your specified criteria.
- **Time Intelligence**: Parses date columns to automatically allocate corresponding **Quarters** (Q1, Q2, Q3, Q4) and **Years**.
- **Automated Pivot Tables**: Automatically configures and generates a Pivot Table based on your cleaned data.

## Prerequisites

- **Microsoft Excel** (Desktop version)
- A macro-enabled workbook (`.xlsm` or `.xlsb`).
- Ensure the **Developer Tab** is enabled in your Excel ribbon.
- This module was built on a sheet with 3 columns: Account name (income type), transaction date, amount of transaction.

## How to Install and Use

1. **Clone or Download** the repository to your local machine:
   ```bash
   git clone https://github.com
   ```
2. Open **Microsoft Excel** and open your target data workbook.
3. Press `Alt + F11` to open the Visual Basic for Applications (VBA) Editor.
4. On the top menu, click **Insert** ➔ **Module**.
5. Copy the code from `IncomeAnalysis.bas` in this repository and paste it into the new module window.
6. **Customize Variables**: Modify the code slightly to match your exact sheet names, date columns, and pivot table field targets.
7. WARNING: Make sure you have the code exactly as you need before running on your sheet as VB code cannot be undone. 
8. Press `F5` to run the macro, or return to Excel, press `Alt + F8`, select the macro, and click **Run**.

## ⚙️ Configuration

Before running the macro for the first time, make sure to update these key variables in the VBA script:

- `TargetSheet`: The name of the worksheet containing your raw data (e.g., `"RawData"`).
- `DateColumn`: The letter of the column containing your dates (e.g., `"C"`).
- `DeleteCriteria`: The condition defining which rows to remove.
- Considering copying your sheet for testing before running on your raw data. 

## Credits

This skill was learned at UTRGV and through WSO. 
