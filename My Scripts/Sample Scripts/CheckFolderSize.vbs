'This Example Demonstrates How To Get Size Of Any Folder.
'It Uses ".Size" Property Of "Folder" Object.

Dim Fso, Msg, FolderObj, FolderPath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FolderPath = InputBox("Enter Full Path Of The Folder : ","Folder Path") 'Get Path Of The Folder
If (Fso.FolderExists(FolderPath)) Then 'Checks Whether Folder Exits At The Specified Path
	Set FolderObj = Fso.GetFolder(FolderPath) 'Returns "Folder" Object
	Msg = "Folder : " & FolderPath & " Uses " & FolderObj.Size & " Bytes" '.Size Property Returns Size Of The Folder In Bytes.
Else 'Folder Doesn't Exit.
	Msg = "Folder : " & FolderPath & " Doesn't Exist."
End If
MsgBox Msg