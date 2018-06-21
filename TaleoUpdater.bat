"C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE" "X:\HRIS\Taleo\Database\Taleo_Updater.accdb" /x START_PAUSE

@REM trying to add TCC to this process; the above has taken so long, TCC isn't running regularly 
@REM because the above isn't done when it tries to kick off
@REM adding it here *should* force it to run after taleo_rpt_imports finishes, every time

@REM sleeps for 5 seconds before continuing
TIMEOUT 5

@REM former Taleo_Run_TCC process
"X:\HRIS\Taleo\TCC\PROD\IMP\TCC Files\main_import.bat"

@REM adding TCC_Results_Check here as well, took 60+ minutes to run the previous two steps as it is now
@REM this only allowed for 50

@REM TIMEOUT 25

@REM former Taleo_Results_Check process - never kicks off for some reason
@REM "\\umh.edu\data\Personnel_Payroll\Salary\HRIS\Automation\TaleoTCCResultsCheck.bat"

EXIT /B %ERRORLEVEL%