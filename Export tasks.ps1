    #<############################################## Export Scheduled Task ##############################################
    $ExportPath = "X:\Users\Sean Ford\Scheduled Tasks"           # Set export path, you need to have access to it.
    $Tasks = Get-ScheduledTask #–TaskPath "\FolderName\" # Set task path (folder name) as it have in Task Scheduler or leave it as "\" if it placed in root.

    # Create path if not exist
    If((Test-Path $ExportPath) -ne $True){New-Item -ItemType directory -Path $ExportPath}

    Foreach($Task in $Tasks) 
    {
        Export-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath | Out-File (Join-Path $ExportPath "$($Task.TaskName).xml")
    }
