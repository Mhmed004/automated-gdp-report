# 🌍 Automated GDP Economic Report Generator

## Overview
A Python automation script that pulls live GDP and economic data for 200+ countries 
directly from the World Bank API and auto-generates a formatted Excel report.

## Features
- Pulls live data from the World Bank API (no manual downloads)
- Analyzes GDP, growth rates, and per capita income for 200+ countries
- Auto-generates a formatted 4-sheet Excel report with charts
- Names output file with today's date automatically
- Runs end-to-end in under 5 seconds

## Technologies
- Python, Pandas, OpenPyXL, wbgapi (World Bank API)

## How to Run
1. Clone the repo
2. Install dependencies: `pip install pandas wbgapi openpyxl`
3. Run: `python sales_report.py`
4. Open the generated Excel file
