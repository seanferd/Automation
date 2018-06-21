##################################################################################
#                   Taleo CSV File Renamer
##################################################################################
#   This script was created to alleviate the Taleo Hour requirement - wherein
#   Taleo would create its CSV files with yyyy_mm_dd_hh. This required the next 
#   step in the process to run in the same hour or everything would fail. I modified
#   the process to remove that requirement, but Taleo would then overwrite its own
#   files if it was re-run on the same day. This script adds the _hh back to the 
#   CSV filename for all files created today. This is hooked into the file:
#   H:\Automation\TaleoTCCResultsCheck.bat - which is the last process that runs
#   in the Taleo 'batch'. 
#
#   This script is the counterpart of file-mover.ps1
#
#                   Written by:
#                       Sean R Ford, UMHC HRIS
#                       fords@health.missouri.edu
#                       12/12/2017
##################################################################################
#   Last Modified: 12/14/2017 - Added flowerbox -srf
#   12/12/2017: File created -srf
##################################################################################

$filePaths = @( "X:\HRIS\Taleo\TCC\PROD\IMP\", 
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\dept_reset_approver_files\results",
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\candidate_imp\results",
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\dept_imp\results",
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\template_imp\results",
                "O:\Salary\HRIS\Taleo\TCC\PROD\IMP\TCC Files\users_imp\results" 
            )
#While testing, this was handy to remove the extra _HH from the files
# foreach ( $path in $filePaths ) {
#     $newPath = $path + '\Archive\' + $year + '\'
#     Get-ChildItem $path -Filter '*.csv' | 
#         Where-Object { $_.CreationTime.Date -eq (Get-Date).Date } | 
#         Rename-Item -NewName { $_.Name -replace '_10', '' }
# }

#pull just the current hour, e.g. 09
$hour = Get-Date -format HH

#What we intend to add to the filename, e.g. _09.csv
$replace = '_' + $hour + '.csv'

foreach ( $path in $filePaths ) {
    #loop through all the CSV files we find that were created today and add the hour to the filename
    Get-ChildItem $path -Filter '*.csv' | 
        Where-Object { $_.CreationTime.Date -eq (Get-Date).Date } | 
        Rename-Item -NewName { $_.Name -replace '.csv', $replace } 
}
