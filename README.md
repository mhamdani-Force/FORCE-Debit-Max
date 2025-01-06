# README for FORCE-Debit Max Script

## Project Overview
This repository contains a Python script designed to process hydrological data stored in Excel files. The primary objective is to analyze and extract maximum discharge values ("Débit Max") based on date and time information from multiple sheets within an Excel file.

## Features
- Interactive file selection through a graphical interface.
- Automatic detection of date, time, and discharge columns.
- Data cleaning and conversion to standard formats.
- Monthly and yearly maximum discharge value extraction.
- Export of results to a new Excel file.

## Prerequisites
- Anaconda (Python 3.x environment).
- Spyder IDE.
- Required Libraries:
  - `pandas`
  - `numpy`
  - `tkinter`

Install the dependencies using:
```bash
conda install pandas numpy
```

## File Requirements
- The input Excel file should have multiple sheets.
- Each sheet must include columns for:
  1. `Date` - containing date values.
  2. `H` - containing time values.
  3. `Débit` - containing discharge values.

## How to Use
1. Open the script in Spyder (part of the Anaconda distribution).
2. Ensure all the code is in the `main()` function for easy execution.
3. Run the script.
4. Select the input Excel file when prompted.
5. The script will process all sheets and identify columns matching the required format.
6. Results containing maximum discharge values per month and year will be saved to a new Excel file. Specify the output path during the save prompt.

## Error Handling
- Missing or invalid columns trigger a descriptive error message.
- Unsupported formats or data types are skipped with warnings.
- Permission errors prevent overwriting open files.

## Example
Input Excel file structure:
```
Sheet1:
Date       | H      | Débit
2024-01-01 | 08:00  | 15.2
2024-01-01 | 12:00  | 18.5
2024-01-02 | 08:00  | 12.7

Sheet2:
Date       | H      | Débit
2024-02-01 | 08:00  | 20.1
2024-02-02 | 12:00  | 21.4
```
Output Excel file:
```
Année | Mois | Date       | H      | Débit
2024  | 1    | 2024-01-01 | 12:00  | 18.5
2024  | 2    | 2024-02-02 | 12:00  | 21.4
```

## Troubleshooting
- Ensure all columns are labeled correctly (Date, H, Débit).
- Close the Excel file before running the script.
- Check that data formats are consistent within the Excel file.

## Author
Mohamed Hamdani

## License
This project is licensed under the MIT License - see the LICENSE file for details.

