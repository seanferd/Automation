##################################################################################
#                   Access Timeout Check
##################################################################################
#   This is a simple passthrough script for our existing batch files. We are having
#   regular issues where a process will hang in the evening and cause our limit
#   of Access processes (seems to be 4) to 'fill up'. Once we hit that limit all of
#   our other nightly processes stack up behind them just waiting for us to hit ok
#   or otherwise close Access. This script is intended to be called from our existing
#   batch files so nothing with Active Batch needs to change. A big part of this decision
#   is that we'd have to change EVERY Task Manager process on the server as well to
#   directly call Powershell, instead of using the existing batch calls. Maybe in the
#   future we can just point AB to this script directly within the Job Steps... the 
#   benefit here would be that AB would know if our process was still running or not -
#   currently it calls the Task Manager which kicks the .bat off and then it thinks
#   the process is done, but it is not. 
#
#   Anyway, we make the call, powershell starts the process and monitors it for a timeout 
#   condition. If it hits that, it will stop its process and send an email about the 
#   failure. This is sort of a temporary solution until we can figure out why some of our 
#   processes are hanging overnight, but should give us more automation and free time 
#   since we won't have to manually stop processes daily and can focus on fixing the 
#   broken ones.
#
#                   Written by:
#                       Sean R Ford, UMHC HRIS
#                       fords@health.missouri.edu
#                       12/18/2017
##################################################################################
#   Last modified: 12/20/2017: Added exitcode return so AB will not always report success -srf
#   12/18/2017: Added flowerbox -srf
#   12/18/2017: File created -srf
##################################################################################

#parameters accepted
#a .bat call should look similar to these:
#powershell.exe .\timeout-check.ps1 "O:\Organizational_Structure\Data\Reporting_Structure_be.mdb" "UPDATE_STRUCTURE"
#powershell.exe .\timeout-check.ps1 "H:\CIA\Covered_Employees.accdb" "compact"
Param(
    [string] $accessPath,
    [string] $macro
)

#the location of the MSAccess executable on the given machine                    
$MsAccess = "C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE"

#concatenated paramaters
$args = ($accessPath) + " /x" + $macro

#special scenario if compact is passed in for some reason
if( $macro -eq "compact") {
    $args = ($accessPath) + " /compact"
}

#starting the Access process and passing in the params
#we needed to call Access explicitly so we could have more than one instance run at once
$proc = start-process $MsAccess $args -PassThru

#this will timeout the process in the event of an error - if we hit the timeout, it populates the timeout variable (-ev timeout)
#setting timeout to 20 minutes; nothing should run that long except Taleo processes, and we don't plan to include them here
$proc | Wait-Process -Timeout 1200 -ev timeout

#if we hit the timeout limit, stop the process and send an email
if( $timeout ) {
    $proc | Stop-Process

    $outlook            = New-Object -ComObject Outlook.Application
    $mail               = $outlook.CreateItem(0)
    $mail.importance    = 2
    $mail.Subject       = "A DB process has timed out (20 minutes)"
    $mail.to            = "fords@health.missouri.edu;ungerj@health.missouri.edu;websterh@health.missouri.edu"
    $mail.HTMLBody      = "A DB process has timed out<br/>"
    $mail.HTMLBody      += "<a href='" + $accessPath + "'>" + $accessPath + "</a><br/><br/>"
    $mail.HTMLBody      += "Macro is: " + $macro
    $mail.send()

    #this should return a 99 exitcode which will pass through to Active Batch
    #we don't care exactly what 99 means, AB only looks at whether it's a zero or non-zero return
    exit(99)
}
