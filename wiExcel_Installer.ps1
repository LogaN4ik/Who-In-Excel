Clear-Host
$wiExcel_dir = "C:\My\Scripts\wiExcel"
$wiExcel_client_dir = "G:\Plan\ОПБиК\2\Скрипты\wiExcel\Client\wiExcel.zip"

#проверить или создать директорию
If(!(test-path $wiExcel_dir)){New-Item -ItemType Directory -Force -Path $wiExcel_dir}

#Копируем клиент с сети на пк
Copy-Item -Path $wiExcel_client_dir -Destination $wiExcel_dir -Force
Expand-Archive -Path $wiExcel_client_dir  -DestinationPath $wiExcel_dir -Force
Remove-item "C:\My\Scripts\wiExcel\wiExcel.zip"


#создаем задачу в планировщик задач
$action = New-ScheduledTaskAction -Execute "C:\My\Scripts\wiExcel\silent_wiExcel.vbs"
$trigger = New-ScheduledTaskTrigger -Daily -At 12:00pm
$task = Register-ScheduledTask -TaskName "wiExcel" -Trigger $trigger -Action $action
$task.Triggers.Repetition.Interval = "PT5M"
$task | Set-ScheduledTask
