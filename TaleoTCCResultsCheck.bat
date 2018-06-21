@REM Seems to timeout because 'Quit' isn't available now in Access
@REM "C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE" "X:\HRIS\Taleo\Database\Taleo_Updater.accdb" /x START_PAUSE_RESULTS

Powershell.exe -NoProfile -ExecutionPolicy Bypass H:\Automation\automation_sean\timeout-check.ps1 "X:\HRIS\Taleo\Database\Taleo_Updater.accdb" "START_PAUSE_RESULTS"

PowerShell -NoProfile -executionpolicy bypass -Command "& X:\HRIS\Taleo\Database\Taleo_Reporting\Taleo_RA_Data\file-rename.ps1" > C:\Users\!ungerj\Desktop\ps-rename.log

EXIT /B %ERRORLEVEL%