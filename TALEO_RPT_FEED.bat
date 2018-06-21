"C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE" "x:\HRIS\Taleo\Database\Taleo_Reporting\Taleorpt.accdb" /x UPDATE_RPT_CHECKS_START

PowerShell -NoProfile -executionpolicy bypass -Command "& X:\HRIS\Taleo\Database\Taleo_Reporting\Taleo_RA_Data\file-mover.ps1" > C:\Users\!ungerj\Desktop\ps-move.log

EXIT /B %ERRORLEVEL%