# Vacant House Finder

## Overview
Vacant House Finder is a Python project designed to process property assessment data, match properties with real estate listings, and generate an Excel report to help identify potential real estate investment leads.

## Features
- Processes property assessment data using Python.
- Filters residential properties based on available property information.
- Calculates a preliminary vacancy indicator score.
- Matches property records with manually downloaded real estate listings.
- Combines multiple Redfin CSV downloads and removes duplicate listings.
- Exports results to Excel for further analysis.
- Includes editable property-condition indicators, such as boarded-up status and overgrown yards.

## Technologies Used
- Python
- pandas
- openpyxl
- Microsoft Excel
- CSV data processing

## How It Works
1. Import property assessment data.
2. Filter and process residential property records.
3. Import real estate listing data downloaded from Redfin.
4. Match listing addresses against property records.
5. Calculate preliminary lead scores.
6. Export results into an Excel workbook for review.

## Important Notes
The vacancy score is a preliminary screening tool, not proof that a property is vacant. Property matches and listing status should be independently verified. Listing results depend on the dates and sources of the downloaded data.

## Purpose
This project demonstrates practical skills in Python programming, data cleaning, data extraction, data matching, Excel automation, and real estate lead analysis.
