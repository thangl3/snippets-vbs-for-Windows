Set objShell = CreateObject("WScript.Shell")
Set objFile = CreateObject("Scripting.FileSystemObject")

Sub Include(strFileName)
	Set objTextFile = objFile.OpenTextFile(strFileName, 1)

	ExecuteGlobal objTextFile.ReadAll

	objTextFile.Close

	Set objTextFile = Nothing
End Sub

Function GetRunningFolderPath
    GetRunningFolderPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
End Function

Include(GetRunningFolderPath & "/TaskScheduleHelper.vbs")
Include(GetRunningFolderPath & "/launcher.vbs")

Call runAsAdministrator

Dim stringEvent, stringCommand, stringArgument, action

action = generateInputBoxForAction()

stringEvent = action(0)
stringCommand = action(1)

If stringCommand <> "" Then
	stringArgument = InputBox("Enter the argument for command", "Input box")

	NS = "http://schemas.microsoft.com/windows/2004/02/mit/task"

	Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0") 

		Set rootElement = xmlDoc.createNode(1, "Task", NS)
			rootElement.setAttribute "version", "1.2"
		xmlDoc.appendChild(rootElement)

			Set triggerElement = xmlDoc.createNode(1, "Triggers", NS)
				Set logonTriggerElement = xmlDoc.createNode(1, "LogonTrigger", NS)
					Set Enabled = xmlDoc.createNode(1, "Enabled", NS)
					Enabled.text = "true"
					logonTriggerElement.appendChild(Enabled)
			triggerElement.appendChild(logonTriggerElement)
		
		rootElement.appendChild(triggerElement)
		rootElement.appendChild(genratePrincipalsXml(NS))
		rootElement.appendChild(generateSettingXml(NS))
		rootElement.appendChild(generateActionXml(stringCommand, stringArgument, NS))

	Set objIntro = xmlDoc.createProcessingInstruction("xml","version='1.0' encoding='UTF-16'")  
	xmlDoc.insertBefore objIntro, xmlDoc.childNodes(0)

	xmlDoc.Save "config.xml"

	objShell.Run "C:\windows\system32\cmd.exe /c SchTasks /CREATE /XML config.xml /TN " & stringEvent, 0, 1

	objFile.DeleteFile("config.xml")

	MsgBox "Create task success", 0, "Success"
End If