# Excel Asset Effectiveness Analysis Toolkit

This repository provides a suite of VBA tools for automating the preparation, analysis, and reporting of asset performance data in Excel. The toolkit includes macros such as **AnalysisSheet**, **CountAlphanumerics**, **RunAnalysis** (master macro) and a custom VBA function **ExtractAllAlphanumeric**, each designed to streamline different aspects of data processing, making asset effectiveness analysis efficient and reliable.

---

## AnalysisSheet: Automated Data Setup and Duplicate Management

**Description**:  
The **AnalysisSheet** macro sets up a worksheet by organizing required columns, identifying duplicates, and adding dynamic calculation fields, preparing the data for asset effectiveness analysis.

### Purpose

AnalysisSheet was created to simplify and standardize data preparation, helping users avoid manual data entry, identify duplicate entries, and prepare key columns for further analysis.

### Problem

For asset performance analysis, preparing data manually is time-consuming and prone to errors, especially when handling large datasets with multiple production metrics, duplicate entries, scattered columns, and missing calculated fields add to the complexity.

### Solution

The AnalysisSheet macro addresses these challenges by:
- Automatically identifying and organizing essential columns.
- Checking for and managing duplicate entries, giving users options to retain or remove duplicates.
- Inserting dynamic fields which reflect the category of analysis being carried out.

### Implementation Details

- **Duplicate Management**: Highlights duplicate entries, allowing users to review or remove them.
- **Dynamic Headers**: Adds calculated fields based on Category, making the macro adaptable to various Categories of analysi.
- **Efficient Column Setup**: Hides non-essential columns, keeping only relevant data visible for analysis.

### Usage Instructions

Refer to the [PrepareAnalysisSheet Usage Guide](docs/PrepareAnalysisSheet_usage_guide.md) for detailed setup and instructions.

---

## ExtractAllAlphanumeric: Custom Function for Text Parsing and Alphanumeric Extraction

**Description**:  
The **ExtractAllAlphanumeric** function extracts alphanumeric sequences from a specified text column, helping to extract each asset for analysis.

### Purpose

ExtractAllAlphanumeric ensures consistent extraction of alphanumeric data which represents each assset for grouping and analysis.

### Problem

Extracting alphanumeric codes from text entries manually is time-consuming and inconsistent.

### Solution

The ExtractAllAlphanumeric function:
- Extracts all alphanumeric sequences from any specified column.
- Outputs these sequences for consistent analysis.

### Usage Instructions

Refer to the [ExtractAllAlphanumeric Usage Guide](docs/ExtractAllAlphanumeric_usage_guide.md) for details.

---
## CountAlphanumerics: Unique Alphanumeric Counting and Aggregation

**Description**:  
The **CountAlphanumerics** macro scans a designated column for unique alphanumeric entries, counts occurrences, and outputs the most frequent entries, providing insight into common asset issues.

### Purpose

CountAlphanumerics aids in identifying recurring asset issues by counting unique alphanumeric codes, helping users quickly understand which assets or issues appear most frequently.

### Problem

Manually counting and summarizing alphanumeric entries, such as asset codes, across large datasets is tedious and prone to inaccuracies.

### Solution

The CountAlphanumerics macro:
- Identifies unique alphanumeric entries in a specified column.
- Counts and aggregates occurrences, outputting the top counts in designated columns for easy review.

### Implementation Details

- **Dictionary-Based Counting**: Uses a dictionary object for efficient counting and aggregation.
- **Customizable Output**: Outputs the most frequent (top 10) entries in descending, providing a quick summary for analysis.

### Usage Instructions

Refer to the [CountAlphanumerics Usage Guide](docs/CountAlphanumerics_usage_guide.md) for detailed instructions.

---

## RunAnalysis: Comprehensive Analysis Execution

**Description**:  
The **RunAnalysis** macro orchestrates the analysis process by running each component macro in sequence, from data preparation to extraction and counting.

### Purpose

RunAnalysis provides a seamless workflow, automatically executing each macro and custom VBA function for comprehensive data processing and analysis.

### Problem

Running each macro individually can be time-consuming, especially when analyzing large datasets.

### Solution

The RunAnalysis macro:
- Combines multiple steps into a single automated sequence.
- Simplifies the analysis workflow, ensuring all macros and VBA function are executed in the correct order.

### Implementation Details

- **Sequential Execution**: Runs AnalysisSheet, ExtractAllAlphanumeric, and CountAlphanumerics, in the proper sequence.
- **Single-Click Execution**: Allows users to trigger the entire analysis process with one macro.

### Usage Instructions

Refer to the [RunAnalysis Usage Guide](docs/RunAnalysis_usage_guide.md) for detailed instructions.
