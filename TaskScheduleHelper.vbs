Function genratePrincipalsXml(NS)
	Set Principals = xmlDoc.createNode(1, "Principals", NS)
		Set Principal = xmlDoc.createNode(1, "Principal", NS)
		Principal.setAttribute "id", "Author"
			Set LogonType = xmlDoc.createNode(1, "LogonType", NS)
			LogonType.text = "S4U"
			Set RunLevel = xmlDoc.createNode(1, "RunLevel", NS)
			RunLevel.text = "HighestAvailable"
		Principal.appendChild(LogonType)
		Principal.appendChild(RunLevel)
	Principals.appendChild(Principal)

	Set genratePrincipalsXml = Principals
End Function

Function generateSettingXml(NS)
	Set Settings = xmlDoc.createNode(1, "Settings", NS)
		Set MultipleInstancesPolicy = xmlDoc.createNode(1, "MultipleInstancesPolicy", NS)
			MultipleInstancesPolicy.text = "IgnoreNew"
		Set DisallowStartIfOnBatteries = xmlDoc.createNode(1, "DisallowStartIfOnBatteries", NS)
			DisallowStartIfOnBatteries.text = "true"
		Set StopIfGoingOnBatteries = xmlDoc.createNode(1, "StopIfGoingOnBatteries", NS)
			StopIfGoingOnBatteries.text = "true"
		Set AllowHardTerminate = xmlDoc.createNode(1, "AllowHardTerminate", NS)
			AllowHardTerminate.text = "true"
		Set StartWhenAvailable = xmlDoc.createNode(1, "StartWhenAvailable", NS)
			StartWhenAvailable.text = "false"
		Set RunOnlyIfNetworkAvailable = xmlDoc.createNode(1, "RunOnlyIfNetworkAvailable", NS)
			RunOnlyIfNetworkAvailable.text = "false"
		Set IdleSettings = xmlDoc.createNode(1, "IdleSettings", NS)
			Set RestartOnIdle = xmlDoc.createNode(1, "RestartOnIdle", NS)
				RestartOnIdle.text = "false"
			Set StopOnIdleEnd = xmlDoc.createNode(1, "StopOnIdleEnd", NS)
				StopOnIdleEnd.text = "true"
		IdleSettings.appendChild(StopOnIdleEnd)
		IdleSettings.appendChild(RestartOnIdle)

		Set AllowStartOnDemand = xmlDoc.createNode(1, "AllowStartOnDemand", NS)
			AllowStartOnDemand.text = "true"
		Set Enabled = xmlDoc.createNode(1, "Enabled", NS)
			Enabled.text = "true"
		Set Hidden = xmlDoc.createNode(1, "Hidden", NS)
			Hidden.text = "false"
		Set RunOnlyIfIdle = xmlDoc.createNode(1, "RunOnlyIfIdle", NS)
			RunOnlyIfIdle.text = "false"
		Set WakeToRun = xmlDoc.createNode(1, "WakeToRun", NS)
			WakeToRun.text = "true"
		Set ExecutionTimeLimit = xmlDoc.createNode(1, "ExecutionTimeLimit", NS)
			ExecutionTimeLimit.text = "P3D"
		Set Priority = xmlDoc.createNode(1, "Priority", NS)
			Priority.text = 7
		Set RestartOnFailure = xmlDoc.createNode(1, "RestartOnFailure", NS)
			Set Interval = xmlDoc.createNode(1, "Interval", NS)
				Interval.text = "PT5M"
			Set Count = xmlDoc.createNode(1, "Count", NS)
				Count.text = 15
		RestartOnFailure.appendChild(Interval)
		RestartOnFailure.appendChild(Count)

	Settings.appendChild(MultipleInstancesPolicy)
	Settings.appendChild(DisallowStartIfOnBatteries)
	Settings.appendChild(StopIfGoingOnBatteries)
	Settings.appendChild(AllowHardTerminate)
	Settings.appendChild(StartWhenAvailable)
	Settings.appendChild(RunOnlyIfNetworkAvailable)
	Settings.appendChild(IdleSettings)
	Settings.appendChild(AllowStartOnDemand)
	Settings.appendChild(Enabled)
	Settings.appendChild(Hidden)
	Settings.appendChild(RunOnlyIfIdle)
	Settings.appendChild(WakeToRun)
	Settings.appendChild(ExecutionTimeLimit)
	Settings.appendChild(Priority)
	Settings.appendChild(RestartOnFailure)

	Set generateSettingXml = Settings
End Function

Function generateActionXml(Command, Arguments, NS)
	Set Actions = xmlDoc.createNode(1, "Actions", NS)
		Actions.setAttribute "Context", "Author"
		Set Exec = xmlDoc.createNode(1, "Exec", NS)

			If (Command <> "") Then
				Set CommandXml = xmlDoc.createNode(1, "Command", NS)
					CommandXml.text = Command

				Exec.appendChild(CommandXml)
			End If

			If (Command <> "") Then
				Set ArgumentsXml = xmlDoc.createNode(1, "Arguments", NS)
					ArgumentsXml.text = Arguments
				Exec.appendChild(ArgumentsXml)
			End If
		
	Actions.appendChild(Exec)

	Set generateActionXml = Actions
End Function

Function generateInputBoxForAction()
	Dim stringEvent, stringCommand

	stringEvent = InputBox("Enter the name of event (required)", "Input box")
	stringCommand = InputBox("Enter the command (required)", "Input box")

	If stringEvent = "" Then
		MsgBox "Please enter the event name", 0, "Error"

		Wscript.Quit
	ElseIf stringCommand = "" Then
		MsgBox "Please enter the command", 0, "Error"

		Wscript.Quit
	End If

	generateInputBoxForAction = Array(stringEvent, stringCommand)
End Function