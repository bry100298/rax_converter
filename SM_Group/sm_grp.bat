@echo off

set "sourceFolder=C:\Users\User\Documents\Project\rax_converter\SM_Group\Inbound"
set "targetFolder=C:\Users\User\Documents\Project\rax_converter\SM_Group\Outbound"

REM Check if source folder exists
if not exist "%sourceFolder%" (
    echo Source folder does not exist: "%sourceFolder%"
    exit /b
)

REM Check if target folder exists, create if not
if not exist "%targetFolder%" (
    mkdir "%targetFolder%"
)

REM Iterate through all XML files in the source folder
for %%f in ("%sourceFolder%\*.xml") do (
    echo Processing file: "%%~nxf"
    
    REM Check if the file exists
    if not exist "%%~f" (
        echo File does not exist: "%%~f"
        exit /b
    )

    REM Convert XML to XLSX using Excel without opening Excel GUI and retain original column names
    powershell -Command "$Excel = New-Object -ComObject Excel.Application; $Workbook = $Excel.Workbooks.Open('%%~dpnf.xlsx'); $Worksheet = $Workbook.Sheets.Item(1); $Range = $Worksheet.UsedRange; $HeaderRange = $Worksheet.Rows.Item(1); for ($col = 1; $col -le $Range.Columns.Count; $col++) { $Header = $HeaderRange.Cells.Item(1, $col).Text; $Header = $Header -replace '/.*?/', ''; $HeaderRange.Cells.Item(1, $col).Value2 = $Header }; $Workbook.Save(); $Excel.Quit();"

    REM Move the XLSX file to the target folder
    move "%%~dpnf.xlsx" "%targetFolder%\%%~nxf.xlsx" > nul
    
    REM Check for errors during the move operation
    if errorlevel 1 (
        echo Failed to move "%%~dpnf.xlsx" to the outbound folder.
        exit /b
    ) else (
        echo Successfully moved "%%~nxf.xlsx" to the outbound folder.
    )
)

echo Conversion complete.
