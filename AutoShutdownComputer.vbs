Set objShell = CreateObject("WScript.Shell")
Set objRegEx = CreateObject("VBScript.RegExp")
Set objFile = CreateObject("Scripting.FileSystemObject")

Sub Include(strFileName)
	Set objTextFile = objFile.OpenTextFile(strFileName, 1)

	ExecuteGlobal objTextFile.ReadAll

	objTextFile.Close

	Set objTextFile = Nothing
End Sub

Include("TaskScheduleXmlHelper.vbs")

objRegEx.Global = True
objRegEx.Pattern = "(\d{2}):(\d{2}):(\d{2})"

Dim strTime, ok
strTime = InputBox("Enter time you want to shutdown(format: hh:mm:ss)", "Input box")

Set myMatches = objRegEx.Execute(strTime)
For Each myMatch in myMatches
  ok = TRUE
Next

Function Format(myDate)
    d = WhatEver(Day(myDate))
    m = WhatEver(Month(myDate))    
    y = Year(myDate)
    Format = y & "-" & m & "-" & d
End Function

Function WhatEver(num)
    If(Len(num)=1) Then
        WhatEver="0"&num
    Else
        WhatEver=num
    End If
End Function

IF ok = TRUE THEN
	Dim currentDate, NS
	currentDate = Format(Now)

	NS = "http://schemas.microsoft.com/windows/2004/02/mit/task"

	Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0") 

		Set rootElement = xmlDoc.createNode(1, "Task", NS)
			rootElement.setAttribute "version", "1.2"
		xmlDoc.appendChild(rootElement)

			Set triggerElement = xmlDoc.createNode(1, "Triggers", NS)
				Set calendarTriggerElement = xmlDoc.createNode(1, "CalendarTrigger", NS)
					Set startBoundary = xmlDoc.createNode(1, "StartBoundary", NS)
					startBoundary.text = currentDate & "T" & strTime
					Set Enabled = xmlDoc.createNode(1, "Enabled", NS)
					Enabled.text = "true"
					Set ScheduleByDay = xmlDoc.createNode(1, "ScheduleByDay", NS)
						Set DaysInterval = xmlDoc.createNode(1, "DaysInterval", NS)
						DaysInterval.text = 1
					ScheduleByDay.appendChild(DaysInterval)

					calendarTriggerElement.appendChild(startBoundary)
					calendarTriggerElement.appendChild(Enabled)
				calendarTriggerElement.appendChild(ScheduleByDay)
			triggerElement.appendChild(calendarTriggerElement)

			rootElement.appendChild(triggerElement)
			rootElement.appendChild(genratePrincipalsXml(NS))
			rootElement.appendChild(generateSettingXml(NS))
			rootElement.appendChild(generateActionXml("C:\windows\system32\cmd.exe", "/k shutdown -s -t 1", NS))

	Set objIntro = xmlDoc.createProcessingInstruction("xml","version='1.0' encoding='UTF-16'")  
	xmlDoc.insertBefore objIntro, xmlDoc.childNodes(0)

	xmlDoc.Save "config.xml"

	objShell.Run "C:\windows\system32\cmd.exe /k SchTasks /CREATE /XML config.xml /TN autoShutdownThis"

	MsgBox "Create auto shutdown with time is " & strTime & " success", 0, "Success"

	objFile.DeleteFile("config.xml")
ELSE
	MsgBox "Fail with your time", 0, "Error"
END IF

