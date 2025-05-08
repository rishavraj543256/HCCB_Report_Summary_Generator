# HCCB Excel Report & Summary Generator

A desktop application that automates the generation of distributor audit reports and summary reports for HCCB (Hindustan Coca-Cola Beverages).

## Features

- **Report Generation**: Automatically processes distributor audit data from Excel files
- **Summary Reports**: Creates consolidated summary reports from individual audit reports
- **User-Friendly Interface**: Simple GUI for selecting input files and generating reports
- **Standalone Application**: Packaged as a Windows executable, no installation required

## Screenshots

![image](https://github.com/user-attachments/assets/270f6d27-1fef-4d6a-a324-b2217da46ab0)


## Usage

1. Download the latest release from the Releases section
2. Extract the ZIP file to a folder of your choice
3. Run the `HCCB Excel Report & Summary Generator.exe` file
4. Select the dump folder containing distributor Excel files
5. Select the plan file with distributor information
6. Click "Generate Report Excel" to create the detailed report
7. Click "Generate Summary Report" to create a summary of all reports

## Requirements

- Windows 7/8/10/11
- No additional software required (all dependencies are bundled)

## Development

### Prerequisites

- Python 3.8+
- Required packages: pandas, openpyxl, tkinter

### Setup

1. Clone the repository
2. Install dependencies: `pip install -r requirements.txt`
3. Run the application: `python main_gui.py`

### Building the Executable

```
pyinstaller hccb_report_generator.spec
```

## License

© 2024 TNBT. All rights reserved

## Author

Developed by Rishav Raj
