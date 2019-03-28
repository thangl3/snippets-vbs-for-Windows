Sub runAsAdministrator
    If Not WScript.Arguments.Named.Exists("elevate") Then
        CreateObject("Shell.Application").ShellExecute WScript.FullName _
        , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
        WScript.Quit
    End If
End Sub

Function isProcessRunning( strComputer, strProcess )
    Dim Process, Processes, strObject

    strObject = "winmgmts://" & strComputer

    Set Processes = GetObject( strObject ).InstancesOf( "win32_process" )

    For Each Process in Processes
        If UCase( Process.name ) = UCase( strProcess ) Then
            isProcessRunning = True
            Exit Function
        Else
            isProcessRunning = False
        End If
    Next

    Set Processes = Nothing
End Function

Function getObjWMIService( agrs )
    Dim objWMIService
    Dim colServices

    Select Case UBound(agrs)
        Case 1
            Set objWMIService = GetObject( "winmgmts://" & agrs(0) & "/root/cimv2" )
            Set colServices = objWMIService.ExecQuery( agrs(1) )
        Case 2
            Set objWMIService = GetObject( "winmgmts://" & agrs(0) & "/root/cimv2" )
            Set colServices = objWMIService.ExecQuery( agrs(1), agrs(2) )
        Case 3
            Set objWMIService = GetObject( "winmgmts://" & agrs(0) & "/root/cimv2" )
            Set colServices = objWMIService.ExecQuery( agrs(1), agrs(2), agrs(3) )
    End Select

    getObjWMIService = Array(colServices)

    Set objWMIService = Nothing
    Set colServices = Nothing
End Function

Sub controlNetwork(flag)
    networkService = getObjWMIService( Array(".", "Select * From Win32_NetworkAdapter where NetConnectionID IS NOT NULL") )

    If "disable" = flag Then
        For Each adapter in networkService(0)
            adapter.disable()
        Next
    End If

    If "enable" = flag Then
        For Each adapter in networkService(0)
            adapter.enable()
        Next
    End If

    WScript.Echo "Done"
End Sub

Function isLaptop()
    Dim batteryService
    Set batteryService = getObjWMIService( Array(".", "Select * From Win32_Battery", , 48) )

    IsLaptop = False

    For Each objItem in batteryService(0)
        IsLaptop = True
    Next

    Set batteryService = Nothing
End Function

Function dblQuote(Str)
    DblQuote = Chr(34) & Str & Chr(34)
End Function

Function runningProcess( strComputer, strProcess )
    Dim Process, strObject

    strObject = "winmgmts://" & strComputer

    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )

        If UCase( Process.name ) = UCase( strProcess ) Then
            
            Exit Function
        End If
    Next
End Function

Sub runProgram(processName, cmd, intStyle, isWait, isDisplayMsg)
    If isProcessRunning ( ".", processName ) = False Then
        Dim objShell
        Set objShell = WScript.CreateObject( "WScript.Shell" )
        objShell.Run cmd, intStyle, isWait

        Set objShell = Nothing
    Else
        If isDisplayMsg <> 0 Then
            msgbox "Maybe," & processName & " is already running", 16, "Notification"
        End If
    End If
End Sub

Sub shell(cmd, intStyle, isWait)
    Dim objShell
    Set objShell = WScript.CreateObject( "WScript.Shell" )

    objShell.Run cmd, intStyle, isWait

    Set objShell = Nothing
End Sub

Function disconnectDrive(strDrive)
    Set objNetwork = WScript.CreateObject("WScript.Network")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If (objFSO.DriveExists(strDrive)) then
        objNetwork.RemoveNetworkDrive strDrive, True, True
        WScript.Echo "Disconnected network drive : " & strDrive
    End If

    Set ObjFSO = Nothing
    Set objNetwork = Nothing
End Function