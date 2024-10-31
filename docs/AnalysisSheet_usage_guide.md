# Usage Guide for AnalysisSheet Macro

## Overview
The **AnalysisSheet** macro is designed to set up and organize an Excel worksheet for asset effectiveness analysis. It automates the identification of required columns, checks for duplicates, and adds dynamic calculation fields, ensuring data is ready for detailed analysis.

## Steps to Use AnalysisSheet

### 1. Open the Workbook and Activate the Target Sheet
- Open the workbook where the **AnalysisSheet** macro is stored.
- Navigate to the worksheet where you want to perform the analysis. Ensure this is the active sheet when you run the macro.

### 2. Ensure Required Columns are Present
To function correctly, the macro looks for the specified columns in the active sheet e.g.
- **Area**
- **Plant**
- **Date**
- **Remarks**
- **Category**
- **Time (hours)**

Make sure these headers are in row 1 of the active sheet.

### 3. Run the AnalysisSheet Macro
- Open the VBA editor by pressing `Alt + F11` and locate the **AnalysisSheet** macro.
- You can also assign the macro to a button in Excel for easy access.
  
### 4. Duplicate Management
If duplicates are found across all the specified columns, a prompt will appear with options:
- **Yes**: Retain only the first instance of each duplicate and remove the others.
- **No**: Leave all duplicates as they are and highlight them for review.
- **Cancel**: Exit the macro without making changes.

### 5. Column Organization and Calculation Fields
After checking for duplicates, the macro will:
- Hide unnecessary columns, leaving only relevant columns visible.
- Insert columns for calculated fields:
  - **Key** column for tracking asset-specific data.
  - **% Loss Time** column, calculated based on the **Time (hours)** column.
  - A dynamic column header, **% of [Category]**, where `[Category]` reflects the Category of asset performance being analyzed or metric being analyzed.

## Additional Notes
- **Data Source Sheet**: The macro references a sheet named "DataSource" in the workbook for metric values. Ensure this sheet exists and contains the relevant data.
- **Dynamic Calculation**: The macro calculates values based on the specific asset performance metric category being analyzed, ensuring adaptability to various asset performance metric.
