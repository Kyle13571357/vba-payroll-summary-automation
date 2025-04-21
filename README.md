# VBA Payroll Summary Automation

This project contains a legacy Excel VBA macro used for automating annual payroll report generation across multiple monthly worksheets.

The script was originally created in a banking environment to streamline repetitive reporting tasks and consolidate payroll-related fields like overtime pay, night shift allowance, and transportation subsidies.

## 📌 Features

- Automatically consolidates key payroll data from each monthly worksheet
- Detects and extracts summary rows labeled "總計" (Total)
- Calculates Month-over-Month (MoM) percentage change
- Applies alternating row coloring and formatted headers
- Generates both high-level and detailed yearly summaries

## 🧠 Technologies

- **Language:** VBA (Visual Basic for Applications)
- **Platform:** Microsoft Excel (Traditional Chinese UI)
- **File Types:** `.xlsm` / `.bas`

## 💼 Real-World Use Case

This macro was actively used in a financial institution to:
- Aggregate 12+ monthly worksheets into a single yearly summary table
- Automate formatting and calculations that were previously done manually
- Save 30–40% of reporting time across operations and general affairs teams

## ⚠️ Note

All data references are anonymized and safe for public viewing. Column headers and worksheet names appear in Traditional Chinese, and inline comments provide English explanations where necessary.

## 📂 Files

- `generate_annual_summary.bas` – Core macro module for building the summary table
- (Optional) `README.md` – This file
