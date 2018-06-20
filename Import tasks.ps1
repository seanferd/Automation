foreach ( $Task in get-childitem "X:\Users\Sean Ford\Scheduled Tasks\*.xml" ){
#echo $Task.BaseName
Register-ScheduledTask -Xml (get-content $Task | out-string) -TaskPath \Automation\ -TaskName $task.BaseName -User "umhs-users\umhc_hris"
}