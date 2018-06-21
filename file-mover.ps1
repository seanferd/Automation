##################################################################################
#                   Taleo CSV File Mover
##################################################################################
#   This script intends to find all CSV files in the paths below (that should be 
#   created by the Taleo process) that were created before today's date. If it 
#   finds any, it will move them to the same path \Archive\YYYY\ directory - if 
#   that directory doesn't exist, it will create it based on the current year. 
#   This should help keep things decluttered and easy to find. This script is 
#   hooked into the first batch file in the Taleo 'batch': 
#   H:\Automation\TALEO_RPT_FEED.bat
#
#   I separated these scripts so the files created today would stay in the directory
#   until tomorrow - I'm not positive if they're being used by anything outside of
#   the Taleo 'batch'
#
#   This script is the counterpart of file-rename.ps1. 
#
#                   Written by:
#                       Sean R Ford, UMHC HRIS
#                       fords@health.missouri.edu
#                       12/14/2017
##################################################################################
#   Last Modified: 12/14/2017 - Added flowerbox, Changed move date to <= today -srf
#   12/14/2017: File created -srf
##################################################################################

$filePaths = @( "X:\HRIS\Taleo\TCC\PROD\IMP\", 
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\dept_reset_approver_files\results",
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\candidate_imp\results",
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\dept_imp\results",
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\template_imp\results",
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\users_imp\results"
            )

#Current year for our Archive directory
$year = Get-Date -format yyyy

foreach ( $path in $filePaths ) {
    #the Archive path we'll use or create if it doesn't exist
    $newPath = $path + '\Archive\' + $year + '\'

    #if the Archive directory \ year doesn't exist, create it before we try to move files to it
    #should only come into play once a year, but at least we don't have to think about this script again
    if( ! (Test-Path $newPath) ) {
        New-Item -ItemType Directory -Force -Path $newPath
    }

    #loop through all the CSV files we find move them to the Archive dir - this was < today, but on subsequent runs we want to go ahead and clear those out
    #so we're left with only the latest file. Since this file runs first, it should always find something and we assume the rest of the Taleo 'batch' will run
    Get-ChildItem $path -Filter '*.csv' | 
        Where-Object { $_.CreationTime.Date -le (Get-Date).Date } | 
        Move-Item -Destination $newPath       
}
