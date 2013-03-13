' Note this script will only work on a computer with Robocopy installed

On Error Resume Next 

' Create Objects
Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

' The location of the pinning shortcuts on your network
serverPath = "\\server\taskbar"
localPath = "C:\taskbar"

' Copy the pinnings onto the local machine
If objFSO.FolderExists(localPath) Then
	strCommand = "RoboCopy " & Chr(34) & serverPath & Chr(34) & " " & Chr(34) & localPath & Chr(34) & " /MIR /XD DfsrPrivate"
	wshShell.run strCommand ,0 ,true
End If

' Open an FSO object at the supplied location
Set folder = objFSO.GetFolder(localPath)
' Open a Shell object at the location
Set objFolder = objShell.Namespace(localPath)

' Iterate through all files
For Each file in folder.Files
	' Set pinnable to True
	pinnable = True

	Set objFile = objFolder.ParseName(Replace(file, localPath & "\", ""))

	' Test if the target exists before pinning
	extArray = split(file,".")
	ext = extArray(ubound(extArray))
	If lcase(ext) = "lnk" Then
		'make shortcut object to be able to test targetpath
		set objShortcut = wshShell.Createshortcut(objFile.Path)
		'confirms that this shortcut is for locally installed software
		If lcase(left(objShortcut.TargetPath,1)) = "c" Then
			'if the target doesnt exist set pinnable to false
			If Not objFSO.FileExists(objShortcut.TargetPath) Then	
				pinnable = False
			End If
		End If
	End If

	If pinnable Then
		' Open a Shell object for the file
		Set fileVerbs = objFile.Verbs
	
		' Iterate through all the verbs
		For Each fileVerb in fileVerbs
			' If the verb is pin to taskbar execute the action
			If Replace(fileVerb.name, "&", "") = "Pin to Taskbar" Then fileVerb.DoIt
		Next
	End If
Next