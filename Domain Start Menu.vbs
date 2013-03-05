' Note this script will only work on a computer with Robocopy installed

On Error Resume Next 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

'Set path to shared startmenu
serverStartMenu = "\\localhost\startmenu"

'Path to local folder you want to create the start menu in
localpath = "C:\startmenu"

'Checks to see if on the network before running the rest of the script
If objFSO.FolderExists(serverStartMenu) Then
	'create string to run robocopy command, add any additional folders you want to exclude after /XD
	'please view robocopy documentation before altering this line
	objCommand = "RoboCopy " & Chr(34) & serverStartMenu & Chr(34) & " " & Chr(34) & localpath & Chr(34) & " /MIR /XD DfsrPrivate"
	
	'run the robocopy command
	wshShell.run objCommand ,0 ,true
	
	'start the search of the local start menu to remove shortcuts for local programs that are not installed
	SearchFolder objFSO.GetFolder(localpath)
End If


Sub SearchFolder(folder)	

	Set objFolder = objFSO.GetFolder(folder.Path)

	'finds extension for each file in folder and if lnk file (shortcut) run test to see if its valid on this computer
	For Each objFile in objFolder.Files
		extArray = split(objFile.Name,".")
		ext = extArray(ubound(extArray))
		If lcase(ext) = "lnk" Then
			'make shortcut object to be able to test targetpath
			set objShortcut = wshShell.Createshortcut(objFile.Path)
			'confirms that this shortcut is for locally installed software
			If lcase(left(objShortcut.TargetPath,1)) = "c" Then
				'if the target doesnt exist remove shortcut from the computer
				If Not objFSO.FileExists(objShortcut.TargetPath) Then
					objFSO.DeleteFile objFile
				End If
			End If
		End If
	Next
	
	'recursively search through all subfolders deleting any empty ones
	For Each subFolder in Folder.SubFolders
		SearchFolder subFolder
		if ((subFolder.files.count = 0) AND (subFolder.subFolders.count = 0)) then
			subFolder.delete 
		end if
	Next
	
End Sub
