# Absence Viewer

A simple HTML/JavaScript application to view and interpret company absence data from Excel files.

## Features

- Upload .xls or .xlsx files containing absence data
- Generates a report showing absence hours by person and calendar week
- Runs entirely in the browser (front-end only)

## Data Format

The Excel file should contain columns including:
- Person Number
- Last Name
- First Name
- Absence start date
- Absence end date
- Absence duration
- Department
- Line Manager
- Start_Date_Duration (hours on start date)
- End_Date_Duration (hours on end date)
- Normal Working Hours (per 5-day week)

## Usage

1. Open `index.html` in a web browser
2. Click the file input and select your .xls or .xlsx file
3. The report will be generated and displayed as a table

## How it works

- Parses the Excel file using SheetJS library
- Creates a composite name from Last Name and First Name
- Calculates absence hours per day based on start/end durations and normal working hours
- Distributes hours across calendar weeks (ISO week numbers)
- Aggregates hours per person per week
- Displays the result in a sortable table

## Dependencies

- SheetJS (xlsx) library loaded via CDN

## Troubleshooting

- Ensure your Excel file has the expected column names
- Dates should be in a recognizable format
- If no data appears, check the browser console for errors
- For large files, processing may take a moment