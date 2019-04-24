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
				Set eventTriggerElement = xmlDoc.createNode(1, "EventTrigger", NS)
					Set Enabled = xmlDoc.createNode(1, "Enabled", NS)
					Enabled.text = "true"
					Set Subscription = xmlDoc.createNode(1, "Subscription", NS)

						Set QueryList  = xmlDoc.createNode(1, "QueryList", "")
							Set Query = xmlDoc.createNode(1, "Query", "")
								Query.setAttribute "Id", "0"
								Query.setAttribute "Path", "System"

								Set SelectElm = xmlDoc.createNode(1, "Select", "")
									SelectElm.setAttribute "Path", "System"
									SelectElm.text = "*[System[Provider[@Name='USER32'] and EventID=1074]]"

							Query.appendChild(SelectElm)
						QueryList.appendChild(Query)
					Subscription.text = QueryList.xml

				eventTriggerElement.appendChild(Enabled)
				eventTriggerElement.appendChild(Subscription)
			triggerElement.appendChild(eventTriggerElement)
		
		rootElement.appendChild(triggerElement)
		rootElement.appendChild(genratePrincipalsXml(NS))
		rootElement.appendChild(generateSettingXml(NS))
		rootElement.appendChild(generateActionXml(stringCommand, stringArgument, NS))

	Set objIntro = xmlDoc.createProcessingInstruction("xml","version='1.0' encoding='UTF-16'")  
	xmlDoc.insertBefore objIntro, xmlDoc.childNodes(0)

	xmlDoc.Save "config.xml"

	objShell.Run "C:\windows\system32\cmd.exe /c SchTasks /CREATE /XML config.xml /TN " & stringEvent, 0, 1

	MsgBox "Create done", 0, "Success"

	objFile.DeleteFile("config.xml")
End If