On Error Resume Next 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

serverStartMenu = "\\greatwood.local\users\startmenu"
localpath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\School StartMenu"

If objFSO.FolderExists(serverStartMenu) Then

	objCommand = "RoboCopy " & Chr(34) & serverStartMenu & Chr(34) & " " & Chr(34) & localpath & Chr(34) & " /MIR /XD DfsrPrivate"

	wshShell.run objCommand ,0 ,true

	SearchFolder objFSO.GetFolder(localpath)
End If


Sub SearchFolder(Folder)	
	Set objFolder = objFSO.GetFolder(Folder.Path)
	Set colFiles = objFolder.Files
	
	For Each objFile in colFiles
		extArray = split(objFile.Name,".")
		ext = extArray(ubound(extArray))
		If lcase(ext) = "lnk" Then
			set objShortcut = wshShell.Createshortcut(objFile.Path)
			If lcase(left(objShortcut.TargetPath,1)) = "c" Then
				If Not objFSO.FileExists(objShortcut.TargetPath) Then
					objFSO.DeleteFile objFile
				End If
			End If
		End If
	Next
	
	For Each Subfolder in Folder.SubFolders
		SearchFolder Subfolder
	Next
	
End Sub
