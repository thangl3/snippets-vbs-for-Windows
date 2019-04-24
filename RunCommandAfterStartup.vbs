Dim commandName, command

commandName = InputBox("Enter the name of command (required)", "Input box")
command = InputBox("Enter the command (required)", "Input box")

If commandName <> "" And command <> "" Then
    Dim userConfigDestination, objShell, WshShell, envAppDATA

    Set objShell = CreateObject("Shell.Application")  
    Set WshShell = CreateObject("WScript.Shell")

    envAPPDATA = WshShell.expandEnvironmentStrings("%APPDATA%")

    userConfigDestination = envAPPDATA & "\Microsoft\Windows\Start Menu\Programs\Startup\"

    Dim File, FSO 

    Set FSO = CreateObject("Scripting.FileSystemObject")

    Set File = FSO.CreateTextFile(commandName & ".bat", True)
    File.WriteLine command & ""

    File.Close

    FSO.MoveFile commandName & ".bat", userConfigDestination
    
    Set File = Nothing

    If Err.Number = 0 Then
        WScript.Echo "Added successfully."
    Else
        WScript.Echo "Have a problem, please try again." 
    End If

    Set FSO = Nothing
End If