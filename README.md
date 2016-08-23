# Formatting-Epicor-Backlog-Report
Reformats Epicor's Sales Order Backlog report into tabulated data

Purpose: Epicor (ERP System) can export a sales order backlog report to Excel, however it is printer-friendly but not spreadsheet-friendly.
This prevents further analysis from being done on the Sales Order Backlog. This repository contains code to include an Excel Add-in that
reformats the report on a new sheet, and makes two SQL queries to the database to get Stock status of the part and Method of Manufacturing
information.
