# Project Title: Generate Excel from DOCX

## Description
This project reads song details from a DOCX file named `cuprins_nou.docx` and generates an Excel file containing the extracted information. The Excel file includes filters on the columns: number, name, range, and theme.

## Requirements
To run this project, you need to install the following dependencies:

- `python-docx`: For reading DOCX files.
- `pandas`: For data manipulation and creating Excel files.
- `openpyxl`: For writing Excel files.

You can install the required packages using pip:

```
pip install -r requirements.txt
```

## Usage
1. Place the `cuprins_nou.docx` file in the appropriate directory.
2. Run the main script:

```
python src/main.py
```

3. The output will be an Excel file generated in the same directory, containing the song details with filters applied.

## Structure
- `src/main.py`: Main script for processing the DOCX file and generating the Excel output.
- `src/utils/__init__.py`: Utility functions for reading and formatting data.
- `requirements.txt`: Lists the required Python packages.
- `README.md`: Documentation for the project.