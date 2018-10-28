'This Example Demonstrates How To First Create A "FileSystemObject" Object, 
'And Then Use The "FolderExists" Method To Check If The Folder Exists.

Dim Fso, Msg, FolderPath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FolderPath = InputBox("Enter Full Path Of The Folder : ","Folder Path") 'Get Path Of The Folder
If (Fso.FolderExists(FolderPath)) Then 'Checks Whether The Specified Folder Exists.
	Msg = "Folder : " & FolderPath & " Exists."
Else
	Msg = "Folder : " & FolderPath & " Doesn't Exist."
End If
MsgBox Msg 'Show The Message