#  Excel-AggReconciler

**Excel-AggReconciler** is a desktop application designed to automate the reconciliation of complex datasets (e.g., Accounting records vs. Vendor statements). 

Unlike simple row-by-row matchers, this tool uses an **Aggregation Algorithm**: it groups multiple transactions by a unique ID (Tax ID, SKU, Contract #), sums them up, and compares the *actual* totals between two files.

###  Why I built this?
Manual reconciliation in Excel using `VLOOKUP` is slow and error-prone, especially when one client has 50 small invoices in one file and 1 big payment in another. This tool solves that problem instantly.

###  Key Features
- **Smart Aggregation:** Automatically groups rows by Key (e.g., BIN/IIN) and calculates the sum before comparing.
- **Flexible Mapping:** Works with any column names – just map them in the UI.
- **Total Verification:** Generates a "Grand Total" row to ensure 100% data integrity.
- **Data Cleaning:** Handles spaces in numbers, text formats, and messy inputs.
- **Modern GUI:** Built with `customtkinter` for a clean, dark-mode experience.

- A must-have tool for Accountants and Financial Analysts!
- ##  Localization / Language
Please note that the **GUI labels** and **internal code comments** are currently in **Russian**, as this tool was originally tailored for the CIS market (Kazakhstan/Russia) and local accounting standards (1C ERP).

*English localization for the interface is planned for future updates.*
---
