:: .csv to .xlsx conversion script

:: Disable the command echo to keep the console output clean
@echo off 
 
:: Enable delayed variable expansion for the script
setlocal enabledelayedexpansion

:: Loop through all CSV files in the current directory
for %%f in (*.csv) do ( 
    :: Store the filename of the first CSV file found in the variable "original_file"
    set "original_file=%%f" 

    :: Store the full path of the first CSV file found in the variable "original_file_path"
    set "original_file_path=%%~ff" 

    :: Exit the loop once the first CSV file is found
    goto :found 
) 

:: Label for the next section of the script
:found 

:: If no CSV file was found, notify the user and exit the script
if not defined original_file ( 
    echo Nessun file CSV trovato nella cartella. 
    pause 
    exit /b 
) 

:: Store the current working directory path in the variable "actual_path"
set "actual_path=%cd%" 

:: Call the WSF script to convert the CSV file to an Excel file (XLSX format)
cscript "%~f0?.wsf" "!original_file_path!" "!actual_path!\!original_file:.csv=.xlsx!" 

:: Open the converted Excel file in Microsoft Excel
start excel.exe %cd%\%original_file:.csv=.xlsx% 

:: Notify the user that the conversion is complete
echo Conversion to .xlsx completed.
::pause 

::--- wsf script --- 
<job>
<script language="VBScript"> 

' Get the source file path from the script arguments
sourceFile = Wscript.Arguments(0) 

' Get the target file path from the script arguments
targetFile = Wscript.Arguments(1) 

' Enable error handling to avoid breaking the script on runtime errors
On Error Resume Next 

' Open the source CSV file for reading
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(sourceFile, 1)

' Read the first line of the CSV file to determine the delimiter type
line = objFile.ReadLine()
objFile.Close

' Try to get an existing instance of Excel
Set tExcel = GetObject(,"Excel.Application") 

' If no Excel instance exists, create a new instance
If Err.Number = 429 Then 
    Set tExcel = CreateObject("Excel.Application") 
End If 

' Keep Excel hidden during the operation
tExcel.Visible = False

' Suppress Excel alert messages
tExcel.DisplayAlerts = False

' Open the CSV file with comma as the delimiter
If InStr(line, ",") > 0 Then
    Set tWorkbook = tExcel.Workbooks.Open(sourceFile, , True, 2)
' Open the CSV file with semicolon as the delimiter
ElseIf InStr(line, ";") > 0 Then
    Set tWorkbook = tExcel.Workbooks.Open(sourceFile, , True, 4)
End If  

' Access the first worksheet in the workbook
Set tWorksheet1 = tWorkbook.Worksheets(1)

' Get the used range of cells in the worksheet
Set tRange = tWorksheet1.UsedRange 

' Auto-fit the columns to adjust their width based on content
tRange.EntireColumn.Autofit() 

' Make the first row bold (usually for headers)
tExcel.Rows(1).Font.Bold = TRUE 

' Save the worksheet as an Excel file in XLSX format (file format 51)
tWorksheet1.SaveAs targetFile, 51 

' Close the workbook without saving changes
tWorkbook.Close False 

' Quit the Excel application
tExcel.Quit() 

</script>
</job>
