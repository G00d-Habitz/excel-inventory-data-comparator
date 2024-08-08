# Excel Data Comparator
![GitHub repo size](https://img.shields.io/github/repo-size/g00d-habitz/excel-inventory-data-comparator)
![GitHub last commit](https://img.shields.io/github/last-commit/g00d-habitz/excel-inventory-data-comparator)

## 1. About
Excel Data Comparator is a Python-based tool designed to compare Excel files containing product information and generate a report highlighting the differences between the inventory from two years. This project automates the process of downloading data, normalizing it, and identifying changes between different versions of product data sheets.

Originally developed as a university assignment, this tool has practical applications in inventory management, product catalog maintenance, and any scenario requiring detailed comparison of spreadsheet data.

## 2. How to Run

### Requirements:
- ![Python](https://img.shields.io/badge/Requires-Python-blue)
- ![PyYAML](https://img.shields.io/badge/Requires-PyYAML-yellow)
- ![Requests](https://img.shields.io/badge/Requires-Requests-orange)
- ![Pandas](https://img.shields.io/badge/Requires-Pandas-brightgreen)
- ![XlsxWriter](https://img.shields.io/badge/Requires-XlsxWriter-purple)

You can install the required libraries with:

```bash
pip install pyyaml requests pandas xlsxwriter
```

### Instructions:

1. **Prepare the Environment:**
   - Ensure the `download` folder exists or `dataSource.yml` contains the correct links to the files.
   - If `dataSource.yml` or `download` folder does not exist, the program will create them.

2. **Download Data (Optional):**
   - The script can download data from URLs specified in `dataSource.yml`.
   - Example `dataSource.yml` structure:
     ```yaml
     2023:
       - https://drive.google.com/uc?export=download&id=1wKpSpTx89dbU3SrKt-3-DqQ62kHOFHSX
     2024:
       - https://drive.google.com/uc?export=download&id=1oYKyW7flL53smo56W9srpFdoUuz6_42x
     ```
   - When prompted, type 'Y' to download the data or 'N' to skip this step.

3. **Compare Data Sheets:**
   - Place the `.xlsx` files to be compared in the `download` folder.
   - Input the names of the two files to compare, separated by a comma (e.g., `2023.xlsx,2024.xlsx`).
   - The script will compare the specified sheets and generate a report.

### Usage:

```bash
python compare.py
```

When prompted:
1. Enter 'Y' to download files from `dataSource.yml`, or 'N' to skip.
2. Enter the names of the Excel files to compare, in the format `older.xlsx,newer.xlsx`.

![Python](https://img.shields.io/badge/Made%20with-Python-blue)
