# Automated Sales Reporting using VBA-Macros
Automated Sales Reporting using VBA Macros streamlines sales data processing in Excel. It imports, cleans, organizes data, and generates dynamic dashboards with a single click. Ideal for sales teams and analysts, it saves time and improves accuracy in regular reporting.

Click here to see the file with completed data: https://tinyurl.com/VBA-SalesReport

## About Data
Data was generated using Faker package on Python.

## Features
- Automated Data Import: Automatically updates sales data from local sources, compatible with .csv and .xlsx filetype.
- Data Cleaning and Preparation: Cleans and organizes raw data for consistent analysis.
- Dynamic Dashboard Generation: Creates a visual dashboard that updates with each data refresh, including various metrics and comprehensive charts.

## Prerequisites
- Microsoft Excel
- Enable Macros: Make sure to enable macros for the VBA scripts to run.

## Usage
Open the master VBA-Macros file *VBA-Challenge-Template.xlsm*, Within this file, there are three worksheets:
- "Dasbor": This is where all the data landed and visualized
- "PivotTable": This sheet is used to analyze the raw data before we visualize them
- "Import Data": This is the sheet when we want to import our very first table using the "Add new data" menu and also when we want to update our latest data on the "Update an existing data" menu.

### When uploading the first table (on the RawData > Upload First folder), you need to upload in the following order:
1. "CustomerData" table
2. "ProductData" table
3. "SalesData" table
<b> After that, you can upload additional data in the "Additional" folder without having to follow the order. </b>

## File List
1. RawData: includes "Upload First" folder and "Additional" folder
2. Color-palette: The palette used for the dashboard
3. VBA-Challenge-Template.xlsm: the work-file template

