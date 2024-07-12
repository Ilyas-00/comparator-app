# Excel Comparison Application with Tkinter

## Description
This Tkinter-based application allows users to compare data between two Excel files and export the comparison results to Excel or a log file.

## Features
- Choose two Excel files for comparison.
- Select columns to compare in each file.
- Display differences between the selected columns in a Treeview.
- Export comparison results to Excel (.xlsx) or a  log file (.LOG).

## Requirements
- Python 3.x (version)
- pandas library (`pip install pandas`)
- tkinter library (usually included with Python installations)

## Usage
1. **Setup:**
   - Ensure Python and the required libraries are installed.

2. **Running the Application:**
   - Execute `main.py` using Python (`python3 main.py`).

3. **Comparing Files:**
   - Click "Choisir les fichiers" to select the Excel files for comparison.
   - Select the columns to compare in each file.
   - Click "Comparer" to perform the comparison and display the results.

4. **Exporting Results:**
   - After comparison, use "Créer fichier log" to create a text log file with the comparison details.
   - Use "Exporter vers Excel" to export the comparison results to an Excel file.
