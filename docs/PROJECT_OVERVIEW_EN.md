# Distribution Manager — Project Overview

Distribution Manager is a Python desktop application designed to support a merchandise allocation workflow across stores, warehouses, and digital channels.

## Business problem

Manual spreadsheet-based allocation requires users to track receipts, products, sizes, available quantities, destinations, and remaining stock across multiple files. This creates repetitive work and increases the risk of over-allocation and inconsistent reports.

## Solution

The application imports and validates Excel files, presents stock by product and size, records allocations by destination, prevents quantities from exceeding available stock, automatically assigns remaining units to the appropriate warehouse, stores the results in SQLite, and exports a formatted Excel report.

## My contribution

- Translated an operational workflow into application rules.
- Designed the local relational data model.
- Implemented Excel validation and aggregation with pandas.
- Built the desktop interface with Tkinter and ttk.
- Added stock controls, status management, filters, and exports.
- Separated persistence and business logic from the UI.
- Added automated tests for critical business rules.
- Prepared a public demo with synthetic data and no confidential information.

## Technology choices

- Python for business logic and automation.
- Tkinter for a lightweight Windows-friendly desktop UI.
- pandas and openpyxl for Excel processing.
- SQLite for a self-contained portfolio demo.
- pytest for repeatable business-rule validation.

## Key learning

The main challenge was not the interface itself, but converting an informal manual process into explicit, testable rules while keeping the tool simple for operational users.

## Author

**Joaquín Baigorria**

- LinkedIn: https://www.linkedin.com/in/joaquin-baigorria-5224b0a0/
- GitHub: https://github.com/joabaigo10
- Repository: https://github.com/joabaigo10/Distribution-Manager
