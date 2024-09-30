# Ukrainian Air Defense Analysis Project

This project analyzes data on missile attacks on Ukraine and the effectiveness of Ukrainian air defense systems. It includes scripts for importing data from Kaggle, cleaning the data, and generating charts and statistics.

## Features

- Import data from a Kaggle dataset about missile attacks on Ukraine
- Clean and process raw data, categorizing missile types
- Generate charts and statistics on missile attacks and air defense effectiveness
- Automatically update charts based on user-selected missile types

## Scripts

### 1. Import and Output Script

This script handles the following tasks:
- Imports data from a Kaggle dataset
- Updates chart data based on user selection
- Generates text descriptions and summary statistics

Key functions:
- `importKaggleDataset()`: Fetches data from Kaggle and imports it into the spreadsheet
- `updateChartData()`: Processes data and updates charts based on selected missile type
- `onEdit(e)`: Triggers chart update when user changes missile type selection

### 2. Clean Data Script

This script is responsible for:
- Checking the structure of the raw data
- Cleaning and processing the raw data
- Classifying missile types

Key functions:
- `checkDataStructure()`: Verifies that the raw data sheet has the expected column structure
- `cleanData()`: Processes raw data, filters out irrelevant entries, and classifies missile types
- `classifyModel(model)`: Categorizes missiles into types (ballistic, cruise, hypersonic, etc.)

## Setup and Usage

1. Set up a Google Sheets document with the following sheets:
   - "Raw Data"
   - "Cleaned Data"
   - "Charts"

2. In the Google Sheets script editor, create two script files and paste the contents of the "Import and Output Script" and "Clean Data Script" respectively.

3. Replace placeholder values:
   - In the Import script, replace `kaggleUsername` and `kaggleKey` with your actual Kaggle credentials
   - Update the `sheetId` in both scripts with your Google Sheets document ID

4. Run the `importKaggleDataset()` function to fetch the latest data from Kaggle.

5. Run the `cleanData()` function to process the raw data.

6. Use the dropdown in cell E1 of the "Charts" sheet to select different missile types and automatically update the charts.

## Note

This project uses sensitive data. Ensure you have the necessary permissions and comply with data usage guidelines when using and sharing the results.
