# finance-reporting-automation-vba

This project automates the creation of financial summary reports using Microsoft Excel VBA.

The solution validates financial transaction data, logs data quality issues, aggregates approved records, and generates an automated management report.

The goal of the project was to simulate a real-world financial reporting process and reduce manual effort in data validation and report preparation.

Features
Reads and processes financial transaction data from Excel
Validates mandatory fields
Checks numeric values and business rules
Logs invalid records automatically
Aggregates data by department
Calculates total transaction amounts
Counts transactions by department
Generates an automated financial report
Creates an error log for data quality monitoring
Validation Rules

The application validates:

Missing departments
Missing cost centers
Non-numeric amounts
Negative or zero amounts
Invalid transaction statuses
Output
Report Sheet

The generated report contains:

Department
Cost Center
Total Amount
Transaction Count
Error Log Sheet

The error log records:

Row number
Issue type
Detailed description
Technologies Used
Microsoft Excel
VBA (Visual Basic for Applications)
Business Value

The solution improves reporting efficiency, increases data quality, and reduces manual validation effort during financial reporting processes.
