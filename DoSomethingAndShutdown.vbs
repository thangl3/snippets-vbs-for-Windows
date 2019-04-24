Sub Include( scriptPath )
    Set fStream = CreateObject("Scripting.FileSystemObject")
    Set oStream = fStream.OpenTextFile(scriptPath, 1)
    
    gScript = oStream.ReadAll()
    oStream.Close

    ExecuteGlobal gScript

    Set oStream = Nothing
End Sub

Function GetRunningFolderPath
    GetRunningFolderPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
End Function

Include(GetRunningFolderPath & "/launcher.vbs")

Call runAsAdministrator

' Replace your command here
shell "C:\windows\system32\cmd.exe /c slack-ooo-personal.cmd say --out --event shutdown", 0, True

' Actual shutdown command <Do not remove>
shell "%comspec% /c shutdown -s -t 10", 0, True

MsgBox "Your action was executed, your computer will be shut down after 10s", 0, "Success"