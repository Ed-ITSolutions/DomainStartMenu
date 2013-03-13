' Create Objects
Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' The location of the pinning shortcuts on your network
location = "C:\pinnings\"

' Open an FSO object at the supplied location
Set folder = objFSO.GetFolder(location)
' Open a Shell object at the location
Set objFolder = objShell.Namespace(location)

' Iterate through all files
For Each file in folder.Files
	' Open a Shell object for the file
	Set objFile = objFolder.ParseName(Replace(file, location, ""))
	Set fileVerbs = objFile.Verbs

	' Iterate through all the verbs
	For Each fileVerb in fileVerbs
		' If the verb is pin to taskbar execute the action
		If Replace(fileVerb.name, "&", "") = "Pin to Taskbar" Then fileVerb.DoIt
	Next
Next