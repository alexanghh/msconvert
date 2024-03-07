$action = New-ScheduledTaskAction -Execute "C:\Users\Docker\AppData\Local\Programs\Python\Python312\python.exe" -Argument "-m flask run --host 0.0.0.0" -WorkingDirectory "\\host.lan\Data\msconvert"
$trigger = New-ScheduledTaskTrigger -AtStartup -RandomDelay 00:00:30

Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "msconvert" -Description "Flask service to convert MS office files to PDF"

