# Usage Guide for RunAnalysis - Master Macro

## Overview
The **RunAnalysis** master macro automates the full analysis process by executing each component macro in sequence. It prepares the data, extracts alphanumeric codes, and counts occurrences, streamlining the analysis into a single automated workflow.

## How the Macro Works
- **Input**: Operates on the active worksheet.
- **Output**: Completes data preparation, extraction, and counting, notifying the user at each step.

### Step-by-Step Instructions

1. **Add the Macro to Your Workbook**
   - Open the **Developer** tab in Excel.
   - Select **Visual Basic** to open the VBA editor.
   - **Insert a New Module**:
     - In the VBA editor, go to the **Insert** menu and select **Module**.
     - Paste the **RunAnalysis** code into the module.

2. **Run the Macro**
   - **Select the worksheet** where the analysis should be conducted.
   - **Run the Macro**:
     - Press `Alt + F8`, select **RunAnalysis**, and click **Run**.

3. The macro will execute each step in sequence:
   - Step 1: Prepares the analysis sheet, setting up necessary columns and managing duplicates.
   - Step 2: Extracts unique alphanumeric codes using a custom function.
   - Step 3: Counts unique entries, displaying the top results.

4. **Final Output**:
   - Upon completion, the macro displays notifications for each completed step, ensuring that all processes were successfully run.

### Troubleshooting
- **Error Message**: Ensure that the active worksheet is correctly prepared and contains the expected columns and data.

---

## Contact
For further assistance, reach out to the repository maintainer.
