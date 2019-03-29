Sub Include( scriptPath )
    Set fStream = CreateObject("Scripting.FileSystemObject")
    Set oStream = fStream.OpenTextFile(scriptPath, 1)
    
    gScript = oStream.ReadAll()
    oStream.Close

    ExecuteGlobal gScript

    Set oStream = Nothing
End Sub

Include ("launcher.vbs")

' Replace your command here
shell """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe""", 0, True

' Actual shutdown command <Do not remove>
shell "%comspec% /c shutdown -s -t 30", 0, True

MsgBox "Your action was executed, your computer will be shut down after 30s", 0, "Success"