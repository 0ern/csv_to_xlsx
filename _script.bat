@echo off 
setlocal enabledelayedexpansion 

:: Trova il primo file .csv nella cartella 
for %%f in (*.csv) do ( 
    set "file_originale=%%f" 
    set "percorso_file_originale=%%~ff" 
    goto :found 
) 

:found 
:: Controlla se il file è stato trovato 
if not defined file_originale ( 
    echo Nessun file CSV trovato nella cartella. 
    pause 
    exit /b 
) 

set "percorso_attuale=%cd%" 
cscript "%~f0?.wsf" "!percorso_file_originale!" "!percorso_attuale!\!file_originale:.csv=.xlsx!" 

:: Chiamare Excel per aprire e convertire il file 
start excel.exe %cd%\%file_originale:.csv=.xlsx% 
::timeout /t 5 /nobreak >nul 

echo Conversione in .xlsx completata. 
::pause 

::--- wsf script --- 
<job><script language="VBScript"> 

sourceFile = Wscript.Arguments(0) 
targetFile = Wscript.Arguments(1) 

On Error Resume Next 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(sourceFile, 1)

line = objFile.ReadLine()
objFile.Close

Set tExcel = GetObject(,"Excel.Application") 

If Err.Number = 429 Then 
    Set tExcel = CreateObject("Excel.Application") 
End If 

tExcel.Visible = False
tExcel.DisplayAlerts = False

If InStr(line, ",") > 0 Then
    Set tWorkbook = tExcel.Workbooks.Open(sourceFile, , True, 2)
ElseIf InStr(line, ";") > 0 Then
    Set tWorkbook = tExcel.Workbooks.Open(sourceFile, , True, 4)
End If  

Set tWorksheet1 = tWorkbook.Worksheets(1)
Set tRange = tWorksheet1.UsedRange 

tRange.EntireColumn.Autofit() 
tExcel.Rows(1).Font.Bold = TRUE 
tWorksheet1.SaveAs targetFile, 51 
tWorkbook.Close False 
tExcel.Quit() 

</script></job>