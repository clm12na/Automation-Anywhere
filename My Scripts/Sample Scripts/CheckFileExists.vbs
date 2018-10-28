'This Example Demonstrates How To First Create A "FileSystemObject" Object, 
'And Then Use The "FileExists" Method To Check If The File Exists.

Dim Fso, Msg, FilePath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FilePath = InputBox("Enter Full Path Of The File : ","File Path") 'Get Path Of The File
If (Fso.FileExists(FilePath)) Then 'Checks Whether File Exists At The Specified Path
	Msg = "File : " & FilePath & " Exists."
Else
	Msg = "File : " & FilePath & " Doesn't Exist."
End If
MsgBox Msg 'Show The Message