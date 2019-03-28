Set objShell = CreateObject("WScript.Shell")
Set objFile = CreateObject("Scripting.FileSystemObject")

Sub Include(strFileName)
	Set objTextFile = objFile.OpenTextFile(strFileName, 1)

	ExecuteGlobal objTextFile.ReadAll

	objTextFile.Close

	Set objTextFile = Nothing
End Sub

Include("TaskScheduleXmlHelper.vbs")

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
	rootElement.appendChild(generateActionXml("C:\windows\system32\cmd.exe", "/k start chrome.exe", NS))

Set objIntro = xmlDoc.createProcessingInstruction("xml","version='1.0' encoding='UTF-16'")  
xmlDoc.insertBefore objIntro, xmlDoc.childNodes(0)

xmlDoc.Save "config.xml"

objShell.Run "C:\windows\system32\cmd.exe /k SchTasks /CREATE /XML config.xml /TN autoSayIamIn"

MsgBox "Create auto say I am in", 0, "Success"

objFile.DeleteFile("config.xml")
