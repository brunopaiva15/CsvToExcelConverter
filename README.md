
# CsvToExcelConverter

A simple tool to convert CSV files to Excel format.

## Features
- Converts CSV files to Excel format (.xlsx).
- Optional table creation for easy data management in Excel.
- Adds a context menu option to convert CSV files directly from File Explorer (Windows only).

## Prerequisites
- .NET Framework (or .NET Core) installed on your Windows machine.
- Modify the Windows Registry to add the context menu option (detailed below).

## Installation
1. Clone or download the repository.
2. Build the solution using Visual Studio or the `dotnet` command line.
3. Copy the generated `CsvToExcelConverter.exe` to your desired directory (e.g., `C:\Temp\`).

## Usage
### From Command Line
Run the application using the command:
```
CsvToExcelConverter.exe "path\to\your\file.csv"
```

### From Windows Context Menu (Recommended)
To add a "Convert to Excel" option in the context menu when right-clicking on a CSV file:
1. Create a `.reg` file (e.g., `AddContextMenu.reg`) with the following content:
    ```
    Windows Registry Editor Version 5.00

    [HKEY_CLASSES_ROOT\*\shell\Convert to Excel File]
    @="Convert to Excel File"

    [HKEY_CLASSES_ROOT\*\shell\Convert to Excel File\command]
    @="\"C:\Temp\CsvToExcelConverter.exe\" \"%1\""
    ```
2. Modify the path in the `.reg` file to point to the location of `CsvToExcelConverter.exe`.
3. Double-click the `.reg` file to add the entries to the Windows Registry.

### Application Workflow
1. Prompts you to create a table in Excel.
2. If creating a table, it asks whether the CSV file contains headers.
3. Reads and processes the CSV data.
4. Saves the output as an Excel file in the same directory as the original CSV.

### Notes
- The application will generate a unique filename if an Excel file with the same name already exists.
- For large files, the application limits processing to 1,048,576 rows (Excel's maximum row limit).

## License
This project is licensed under the MIT License. See `LICENSE` for details.
