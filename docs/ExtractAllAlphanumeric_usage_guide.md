# Usage Guide for ExtractAllAlphanumeric Function

## Overview
The **ExtractAllAlphanumeric** function is a custom VBA function designed to parse text strings and extract alphanumeric sequences containing both letters and numbers. This is helpful for identifying asset codes within text.

# How to Add ExtractAllAlphanumeric to Your Workbook

### Step-by-Step Instructions

1. **Enable the Developer Tab** (if not already enabled):
   - Go to **File > Options > Customize Ribbon**.
   - Check **Developer** in the right-hand list, then click **OK**.

2. **Open the VBA Editor**:
   - Click on the **Developer** tab in Excel.
   - Select **Visual Basic** from the ribbon to open the VBA editor.

3. **Insert a New Module**:
   - In the VBA editor, go to the **Insert** menu and select **Module**.
   - A new blank module will open in the editor.

4. **Paste the Code**:
   - Copy the **ExtractAllAlphanumeric** code from the `ExtractAllAlphanumeric.bas` file in the `src` folder.
   - Paste it into the new module in the VBA editor.

5. **Close the VBA Editor**:
   - Once the code is pasted, close the VBA editor. The function is now ready for use in your Excel workbook.


### Example
| Description                           | Extracted Codes      |
|---------------------------------------|-----------------------|
| Sensor 12AB requires maintenance      | 12AB                 |
| Issue with 34CD saftety valve and 78EF pipe   | 34CD, 78EF           |

### Troubleshooting
- Ensure the input contains at least one code with both letters and numbers for the function to return results.

---

## Contact
For additional help, feel free to reach out to the repository maintainer
