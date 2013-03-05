' Note this script will only work on a computer with Robocopy installed

On Error Resume Next 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

'set the location of the network startmenu and where you would like to create the local startmenu
'locations are set in the array so you can redirect different users to different startmenus
'Format is - "path of startmenu to copy" ; "path of local startmenu"

locationsArray = Array	(	"\\localhost\startmenu\pupils;C:\startmenu\pupils",_
				"\\localhost\startmenu\staff;C:\startmenu\staff"_
			)

'iterate though each item in locationsArray
for each item in locationsArray
	
	'create var for start menu locations
	paths = split(item,";")
	serverStartMenu = paths(0)
	localpath = paths(1)
	
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
next

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
