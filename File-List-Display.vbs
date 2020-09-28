Option Explicit

Dim dir
Dim args
Dim objFileSys
Dim objFile
Dim objFolder
Dim objOutputTextStream
Dim FSO

dir = Wscript.Arguments(0)
args = Wscript.Arguments.Count
 
'File system object
Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
'Create log file
Set objOutputTextStream = objFileSys.OpenTextFile("log.txt", 2, True)

Set FSO = CreateObject("Scripting.FileSystemObject")
ShowSubFolders FSO.GetFolder(dir)
Set FSO = Nothing

'Close TextStream
objOutputTextStream.Close
 
Set objOutputTextStream = Nothing
Set objFolder  = Nothing
Set objFileSys = Nothing 

WScript.Echo "Done!" & vbCrlf


Sub ShowSubFolders(Folder)
	Dim File
	Dim Fname
	Dim Subfolder
    For Each File in Folder.Files
        Fname = File.name
		Wscript.Echo Fname
		objOutputTextStream.WriteLine objFile.Name
    Next
 
    For Each Subfolder in Folder.SubFolders
        ShowSubFolders Subfolder
    Next
End Sub
