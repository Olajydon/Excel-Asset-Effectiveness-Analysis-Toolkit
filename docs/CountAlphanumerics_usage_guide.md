# Usage Guide for CountAlphanumerics Macro

## Overview
The **CountAlphanumerics** macro is designed to scan a specified column in Excel, count unique alphanumeric entries, and output the most frequent entries. This is especially useful for identifying recurring asset codes or tags in a dataset.

## How the Macro Works
- **Input**: Scans the `Key` column (defaulted to **AZ2:AZ100**).
- **Output**: Displays the top 10 most frequent unique entries in columns **BF** (Assets) and **BG** (Counts).

### Example Use Case
If you have a column listing asset issues by code (e.g., "611TF201," "FE201"), the macro counts each unique entry, treating uppercase and lowercase versions as identical, and outputs the frequency of each asset code.

## Steps to Use CountAlphanumerics

### 1. Add the Macro to Your Workbook
1. Go to the **Developer** tab in Excel.
2. Select **Visual Basic** to open the VBA editor.
3. **Insert a New Module**:
   - In the VBA editor, go to the **Insert** menu and select **Module**.
   - Paste the **CountAlphanumerics** code into the module.

### 2. Run the Macro
1. **Select the worksheet** where you want to count the alphanumeric codes.
2. **Run the Macro**:
   - Press `Alt + F8`, select **CountAlphanumerics**, and click **Run**.
   
3. The macro will populate columns **BF** and **BG** with the top unique alphanumeric entries and their counts, respectively.

### Example Output

| Key Column (AZ)     | Asset (BF) | Count (BG) |
|---------------------|------------|------------|
| Fe201               | FE201      | 15         |
| FE201               |            |            |
| Fe202               | FE202      | 10         |
| ...                 | ...        | ...        |

### Additional Notes
- **Case Sensitivity**: The macro automatically treats "Fe201" and "FE201" as identical, ensuring consistent counts.
- **Output Limit**: By default, the macro displays the top 10 most frequent entries.

---

## Troubleshooting
- **Empty Output**: Ensure that the `Key` column (default **AZ2:AZ100**) contains values for the macro to analyze.
- **Incorrect Count**: Verify that asset codes are separated by commas in the `Key` column to allow accurate counting.

---

## Contact
For additional help, feel free to reach out to the repository maintainers.
