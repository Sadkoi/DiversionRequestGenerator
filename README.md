# TC Electric - Diversion Request Generator

## Description
This script is designed to automatically generate diversion requests for the implementation of Communication-Based Train Control (CBTC) on the Crosstown Line (G Line) for New York City Transit. This program was written during my time with TC Electric while working on contract S-48012. The script reads data from an Excel spreadsheet and uses a template to generate diversion requests based on specific criteria.

## Features
- Automatic generation of diversion requests based on provided data in an Excel spreadsheet.
- Customizable user input for contract code, date submitted, starting numbers for different diversion types, etc.
- Utilizes OpenPyXL library to interact with Excel files and manipulate cell data.
- Generates output files in Excel format with generated diversion requests.

## User Input
Before running the script, the user needs to modify the following parameters in the script:

```python
DIVERSION_REQUEST_STARTING_NUMBER = 42
CONTRACT_CODE = "S-48012"
DATE_SUBMITTED = "6-7-23" # USE MM-DD-YY FORMAT
STARTING_1A = 1
STARTING_3A = 1
STARTING_6A = 1
STARTING_7A = 1
STARTING_5A = 1
STARTING_1C = 1
STARTING_3C = 1
```

## Prerequisites
- Python 3.11 or higher
- OpenPyXL library `pip install openpyxl`

## Usage
1. Place the input data in the Excel file `Planned_Work.xlsx` following the required format.
2. Place the diversion request template in the Excel file `DivRequest.xlsx` following the required format.
3. Modify the user input parameters in the script as needed.
4. Run the script.
5. The generated diversion requests will be saved in the Requests directory.

## Data Format
The script expects specific column headers in the Excel file. These headers include "Week", "Diversion", "Section", "Start Date", "End Date", and "Diversion Limits." The script will automatically find and assign column coordinates for these headers.

## Notes
- Ensure that the input Excel file (Planned_Work.xlsx) is in the same directory as the script.
- The script assumes specific station names on the G Line for diversion requests. Modify the stops_g_line list if necessary.

## Future Development & Limitations
- The script will be updated to be able to fill out multiple different versions of spreadsheet templates
- Currently, the openpyxl library does not support the copying of shapes and other graphics
