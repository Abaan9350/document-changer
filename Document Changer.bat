@echo off
REM Use PowerShell to open a file dialog and capture the selected file path
for /f "delims=" %%i in ('powershell -command "Add-Type -AssemblyName System.Windows.Forms; $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.InitialDirectory='C:\Users\User\Downloads'; $ofd.Filter='Word Documents (*.docx)|*.docx'; if($ofd.ShowDialog() -eq 'OK'){ $ofd.FileName }"') do set "inputFile=%%i"

if "%inputFile%"=="" (
    echo No file selected.
    pause
    exit /b
)

REM Call the Python script with the selected file path
python "C:\Users\User\Desktop\Task Automations\Python Files\Document_changer.py" "%inputFile%"
pause
