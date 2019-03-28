Sub Include(strFileName)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(strFileName, 1)

	ExecuteGlobal objTextFile.ReadAll

	objTextFile.Close

	Set objFSO = Nothing
	Set objTextFile = Nothing
End Sub

Include("launcher.vbs")