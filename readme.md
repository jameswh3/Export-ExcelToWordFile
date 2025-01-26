# Overview

This script will read columns and rows from an Excel file and create documents for each row, using the value of a specified title column as the file name and creating headings based on each column.

# PowerShell Requirements

*   [Windows PowerShell 7.0 or higher](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4)
* Excel
* Word
* [ImportExcel Module](https://www.powershellgallery.com/packages/ImportExcel)

# Usage

To use the script, run the `Export-ExcelToWordFile` function with the following parameters:

* `-ExcelFilePath`: The path to the Excel file.
* `-OutputDirectory`: The directory where the Word documents will be saved.
* `-TitleColumn`: The column in the Excel file that will be used as the title for each Word document.
* `-OverwriteExistingFile`: A switch parameter to overwrite existing files with the same name.

Example with `-OverwriteExistingFile` switch:
```ps1
Export-ExcelToWordFile -ExcelFilePath "C:\path\to\your\file.xlsx" `
    -OutputDirectory "C:\path\to\output\directory" `
    -TitleColumn "TitleColumnName" `
    -OverwriteExistingFile
```
Example without `-OverwriteExistingFile` switch:
```ps1
Export-ExcelToWordFile -ExcelFilePath "C:\path\to\your\file.xlsx" `
    -OutputDirectory "C:\path\to\output\directory" `
    -TitleColumn "TitleColumnName"
```