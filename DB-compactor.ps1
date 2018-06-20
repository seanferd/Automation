##################################################################################
#                   Access Database Compactor
##################################################################################
#   This script takes a list of fully qualified database files and compacts them.
#   It looks for a lock file (ldb or laccdb), if it finds one it skips and logs it. 
#   It also has been having an issue with folder permissions not allowing a Compact 
#   operation and will sit at a message box indefinitely. To get around this, we invoke 
#   a 10 minute timeout on the call - if it's not done in that time, we stop the process, 
#   log it and move on. NOTE: if a legitimate compact operation takes longer than 10 minutes, 
#   it will also be killed. I have not experienced this in my testing, but it's possible.
#   It checks for a running Outlook process, it will attempt to start one if nothing is found 
#   (we want Outlook running on the server at all times anyway)
#   Finally, it emails Tiers 2 & 3 the results.
#
#                   Written by:
#                       Sean R Ford, UMHC HRIS
#                       fords@health.missouri.edu
#                       12/04/2017
##################################################################################
#                   Calling this script
#   This script can be called with an optional Switch, -noMail 
#
#   Examples:
#       powershell.exe DB-compactor.ps1 -noMail
#
#   Depending on system settings, it may need to be called with the -ExecutionPolicy Bypass switch
#       powershell.exe -ExecutionPolicy Bypass .\DB-compactor.ps1 -noMail
#
#   Or finally, called without the switch that will send email upon completion
#       powershell.exe DB-compactor.ps1
##################################################################################
#   Last Modified:
#       06/18/2018 - Added formatting for Write-Host output when noMail is used, as well as lock owner 
#            checking (if umhc_hris owns it, we'll try to compact anyway). When this script kills a 
#            process, it often leaves a lock file out there even though it's not really locked -srf
#       06/06/2018 - Added message body to terminal output if -noMail is used -srf
#       06/05/2018 - Fixed a bug where paths containing spaces in the DB array would not open -srf
#       06/04/2018 - Added -noMail switch for testing without sending email -srf
#       12/29/2017 - Added Hard to Fill DB to the list -srf
#       12/18/2017 - Added Database*mdb check and removal -srf
#       12/06/2017 - Added flowerbox -srf
#       12/05/2017 - Added lock checking -srf
#       12/04/2017 - File created -srf
##################################################################################

#input parameter / switches
Param(
    [switch] $noMail
)

#if the switch is used correctly, attempt to tell the user
if( $noMail ) {
    Write-Host "Sending of email disabled..."
}

#powershell requires function definitions BEFORE they're called
Function checkForLock( $dBPath ) {
    #get the last thing we find - we know we will find at least one ., but we only care about the last one for the file extension
    $fileExtension = $dBPath.Split(".")[-1]

    #everything else, except for the file extension
    $noExtension = $dbPath.Substring(0,$dbPath.LastIndexOf('.'))

    #mdb files required the else condition, lmdb was not a valid extension
    if( $fileExtension -eq "accdb" ) {
        $lockFile = $noExtension + ".l" + $fileExtension
    } else {
        $lockFile = $noExtension + ".ldb"
    }

    #Test-Path checks if a file exists or not
    if( Test-Path $lockFile ) {
        #tried this inline with the lockfile check above, but it threw an error as the file was not found
        #lock found, if it's not owned by us we have a genuine lock
        if( (Get-Acl $lockFile).Owner -ne "UMHS-USERS\umhc_hris" ) {
            return $true
        #we found a lock, but we own it so we'll try to compact anyway
        } else {
            return $false
        }
    #No lock found
    } else {
        return $false
    }
}

#array of database locations to compact - look into pulling this from a CSV?
$dbsToCompact = @(  "X:\HRIS\Taleo\Database\Taleo_Reporting\Taleorpt.accdb",
                    "H:\ID_Works\ID_Reporting.accdb",
                    "H:\Variable_Hours_db\Variable_Hours_db_Checker.accdb",
                    "X:\HRIS\Mosby\MosbyUpdater.accdb",
                    "H:\Non-Employees\Non_Employees.accdb",
                    "H:\CIA\Covered_Employees.accdb",
                    "X:\Census\New_Employee_Census.accdb",
                    "X:\HRIS\Pos_Mgr_Rpts_db\RRG_Active_Positions_Unfilled\RRG_Reports.accdb",
                    "X:\HRIS\Taleo\Database\Taleo_Updater.accdb",
                    "H:\Employee_Relations\Employee_Relations_updater.accdb",
                    "H:\Position\vacated_positions.accdb",
                    "X:\Users\Jason_Miller\Development\Action_Reasons\Action_Reasons.accdb",
                    "H:\Navex\Navex.accdb",
                    "H:\Employee_Payroll_Deduction\employee_payroll_deductions.accdb",
                    "O:\Organizational_Structure\Data\reporting_structure_be.mdb",
                    "H:\OrgStructure\Reporting_Structure.accdb",
                    "X:\HRIS\Taleo\Database\Taleo_Reporting\Requisition_dashboard.accdb",
                    "X:\HRIS\PI_Process_Improvement_Updater\PI.accdb",
                    "X:\HRIS\Reports\TERMS_HIRS_REHS_NOTIFIER\TERMS_HIRS_REHS_NOTIFIER.accdb",
                    "X:\HRIS\MUHC_SOM_Employees\MUHC_SOM_Employees.accdb",
                    "X:\Users\Jason_Miller\Development\Employee_Changes\Employee_Changes.accdb",
                    "X:\HRIS\Tickler_System\HR_Tickler_System.accdb",
                    "H:\Acquisitions\Affiliates_db.accdb",
                    "X:\Users\Jason_Miller\Development\Surveys\NH_Exit_Surveys.mdb",
                    "X:\HRIS\Halogen\Databases\Halogen_HRIS_Connect.accdb",
                    "H:\Sam.gov\SAM_db.accdb",
                    "X:\HRIS\Competencies_CED_db\CED_Competencies.accdb",
                    "H:\Employee_Survey\Employee_Engagement.accdb",
                    "X:\HRIS\LeadershipRetreat_Updater\LeadershipRetreatPull.accdb",
                    "X:\HRIS\GN_Proact_Positions\GN_Proact_Positions.accdb",
                    "H:\Benchmarks\RN Metrics Reports\Benchmarking(TEST).accdb",
                    "H:\SLRP\Hard To Fill\Hard_To_Fill_Dashboard.accdb"
                )

#the location of the MSAccess executable on the given machine                    
$MsAccess = "C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE"

#loops the array of DBs, begins the compact process and waits for 10 minutes (600 seconds)
#if we hit the timeout limit, it will kill the process it spawned
#skips the DB if we find a lock file
foreach( $path in $dbsToCompact ) {

    #if we don't find a lockfile, proceed
    if( -Not ( checkForLock( $path )) ) {

        #spawns Access and passes the arguments, passthru allows us to watch the process for completion
        #this way, if one finishes we can immediately begin the next DB instead of waiting for an arbitrary timeout
        $proc = start-process $MsAccess -ArgumentList "`"$path`" /compact" -PassThru

        #this will timeout the process in the event of an error - if we hit the timeout, it populates the timeout variable (-ev timeout)
        $proc | Wait-Process -Timeout 600 -ev timeout

        #if we hit the timeout limit, stop the process - intended to allow the script to continue on unknow errors
        #note that this can timeout on a legitimate compact job if it takes too long (not recommended to set it too low)
        if( $timeout ) {
            $proc | Stop-Process
            $failures += @( $path )

            #strips the filename off of the variable
            $dbPath = $path.Substring( 0, $path.LastIndexOf( '\' ) )
            $nowtime = Get-Date

            #if we find a Database*mdb file leftover from our process we'll delete it
            #it sometimes happens when we have to timeout
            foreach( $found in Get-ChildItem $dbPath -Filter 'Database*mdb' ) {
                
                $createtime = $found.CreationTime

                #if the create time of the file is less than an hour old we assume we created it 
                #since this runs at 1am it's a safe assumption
                if( ( $nowtime - $createtime ).totalhours -lt 1 ) {
                    $fullPath = $dbPath + "\" + $found
                    Remove-Item -Path $fullPath
                } 
            }
        }
    } else {
            $locks += @( $path )
            continue
    }
}

    #standard mail options
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.importance = 2
    $mail.Subject = "DB Compactor"
    $mail.to = "fords@health.missouri.edu;ungerj@health.missouri.edu;websterh@health.missouri.edu"

    #if we have failures trying to compact, create a list of the paths that didn't work
    if( $failures.Count -gt 0 ) {

        $mail.HTMLBody = "The following DBs were not able to be compacted, for reasons not related to locks: <br/><br/> `n"
        $mail.HTMLBody += "This failure seems to be related to directory permissions where the DB exists <br/><br/> `n `n" 
        foreach( $db in $failures ) {
            $mail.HTMLBody += "<a href='" + $db + "'>" + $db + "</a><br/><br/>`n"
        }
    } 

    if( $locks.Count -gt 0 ) {
        $mail.HTMLBody += "The following DBs were not able to be compacted because they are locked: <br/><br/>`n `n" 
        foreach( $db in $locks ) {
            $mail.HTMLBody += "<a href='" + $db + "'>" + $db + "</a><br/><br/>`n"
        }
    } 

    if( $failures.Count -eq 0 -and $locks.Count -eq 0) {
        #if we get here, the script ran without finding locks or failures
        $mail.HTMLBody = "The script ran successfully and did not find any locks or failures"
    }

#if the -noMail switch is set, dump the list of failures and exit
#otherwise, send email
if( $noMail ) {
    Write-Host $mail.HTMLBody
    Exit
} else {

    #if outlook isn't running, start it
    if( -Not ( Get-Process -Name "outlook" -ErrorAction SilentlyContinue ) ) {
        #hopefully starts outlook in the event that it crashed or wasn't running
        Start-Process outlook
    }
        #finally send the email we created above
        $mail.Send()
}