# Insurance Statutory Filings Parser

## Overview

This Python script provides a robust and automated solution for parsing complex insurance statutory filings exported from S&P Capital IQ. It is specifically designed to handle the unique XML formats of **Page 19 (Exhibit of Premiums and Losses)** and **Schedule P (Part 1 - Summary)** reports.

The script intelligently identifies the report type within each file, routes it to the correct custom parser, extracts key data points, and aggregates the results from multiple files into a single, clean, multi-tab Excel spreadsheet ready for analysis.

## Key Features

-   **Automatic Report Identification**: Intelligently peeks inside each XML file to determine if it's a Page 19 or Schedule P report.
-   **Batch Processing**: Processes all `.xml` files placed in the `./Inputs/` directory in a single run.
-   **Handles Complex Layouts**:
    -   Correctly parses the multi-table structure of Schedule P reports, extracting data from separate blocks for Premiums, Claims, and Losses.
    -   Handles both country-wide ("GRAND_TOTAL") and state-specific Page 19 reports, creating a `State` column for granularity.
-   **Robust Data Anchoring**: Uses the numbered column headers (e.g., `1`, `25`, `26`) as stable anchors to find data, making the script resilient to changes in text headers or column positions.
-   **Data Cleaning & Formatting**: Cleans numeric data by removing commas and handling non-numeric placeholders.
-   **Deduplication**: Automatically removes duplicate entries in the Page 19 data based on a unique combination of `NAIC`, `YEAR`, `State`, and `LOB`.
-   **Organized Output**: Consolidates all parsed data into a single, timestamped Excel (`.xlsx`) file in the `./Output/` directory, with Page 19 and Schedule P data neatly separated into their own tabs.

## How to Use

### 1. Setup

-   **Clone the repository** (or download the script into a new project folder).
-   **Install dependencies**: This script requires `pandas`, `lxml`, and `xlsxwriter`.
    ```bash
    pip install pandas lxml xlsxwriter
    ```
-   **Create Directories**: In the same directory as the script, create two folders:
    -   `Inputs`
    -   `Output`

### 2. Data Source

The script is designed to work with XML files exported from the S&P Capital IQ platform.

-   **Data Location**: You can download the necessary filings from the "Select-A-Page" feature at the following link:
    [https://www.capitaliq.spglobal.com/web/client?auth=inherit#insurance/select-a-page](https://www.capitaliq.spglobal.com/web/client?auth=inherit#insurance/select-a-page)
-   **Export Format**: When exporting, ensure you select the **XML** format.

### 3. Running the Script

1.  Place all your exported `.xml` files (both Page 19 and Schedule P) into the `./Inputs/` folder.
2.  Run the script from your terminal:
    ```bash
    python your_script_name.py
    ```
3.  The script will process all files, log its progress in the terminal, and generate a single Excel file with a timestamp (e.g., `Combined_Output_20250810_183000.xlsx`) in the `./Output/` folder.

## Output Format

The final Excel workbook will contain two sheets:

### `Page 19 Data` Sheet

| YEAR | Compan_Name | NAIC | State | Liability | LOB | GWP | EP | LOSSES_INCURRED |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| ... | ... | ... | ... | ... | ... | ... | ... | ... |

### `Schedule P Data` Sheet

| REPORT_YEAR | Company_Name | NAIC | YEAR | EP | LOSSES_INC | CLAIMS |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| ... | ... | ... | ... | ... | ... | ... |

