# Customer Data to Kiosk Code Converter

## Overview
This Python-based tool automates the process of converting customer data from Excel files into machine-readable codes for kiosks. It features a user-friendly GUI for easy interaction, dynamic data processing, and flexible export options, making it ideal for streamlining MGT processes in customer data management.

## Key Features
- **GUI for Easy Interaction**: Utilizes `tkinter` for a graphical user interface, allowing users to seamlessly select files and choose export options.
- **Advanced Data Processing**: Employs `openpyxl` to manipulate Excel files, adding new columns and transforming data based on predefined criteria.
- **Flexible Export Options**: Offers the user the choice to export processed data into a single file or multiple segmented files based on specific attributes.

## Prerequisites
Before you can run this script, ensure the following prerequisites are met:
- **Python Installation**: Python 3.x must be installed on your system. Download it from [python.org](https://python.org) and make sure to add Python to your PATH.
- **Required Libraries**:
  - `openpyxl` for handling Excel files.
  - `tkinter` for the GUI (usually included with Python).

You can install `openpyxl` using pip:
```bash
## Installation
- **Obtain the Script: **Download the Python script from the provided link or the shared location.
- **Save the Script:** Ensure the script is saved in a known directory on your computer, avoiding cloud-synced folders to prevent sync issues.
## Usage
To run the script, navigate to the directory where the script is saved and execute the following command in the command prompt or terminal:

**bash**
- Copy code
- python MGT.py

Follow the on-screen prompts to:
- Select the master Excel file.
- Select the MGT Prefix Codes Excel file.
- Choose the export option ('ALL' for a single file or 'BY SEGMENT' for separate files).

## Script Functionality
**Components**
- ExcelProcessor Class: Manages the loading, processing, and exporting of Excel workbooks.
- load_workbooks(): Loads the main and lookup Excel files.
- process_data(): Inserts columns, processes data by extracting and transforming specified columns, and organizes data into segments.
- export_data(): Exports data based on the user's choice of format.

**Utility Functions:**
- **select_excel_file():** Opens a file dialog for file selection.
**Detailed Workflow**
- **Initialize Processor: **Initializes with paths to two Excel files.
- **Load Workbooks: **Reads the Excel files.
- **Prepare Data:** Prepares lookup data from the lookup workbook.
- **Process Data:** Processes data based on predefined rules and organizes by segment.
- **Export Data:** Provides options for data export, either consolidated or segmented.

Notes
Ensure that the Python version is compatible (3.x recommended).
The script is designed to work across Windows, macOS, and Linux operating systems.
