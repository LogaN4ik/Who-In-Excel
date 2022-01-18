command = "powershell.exe -nologo -noninteractive -command C:\My\Scripts\wiExcel\wiExcel.ps1"
set shell = CreateObject("WScript.Shell")
shell.Run command,0, false
