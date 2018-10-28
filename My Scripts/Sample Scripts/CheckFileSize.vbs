'This Example Demonstrates How To Get Size Of Any File.
'It Uses ".Size" Property Of "File" Object.

Dim Fso, Msg, FileObj, FilePath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FilePath = InputBox("Enter Full Path Of The File : ","File Path") 'Get Path Of The File
If (Fso.FileExists(FilePath)) Then 'Checks Whether File Exits At The Specified Path
	Set FileObj = Fso.GetFile(FilePath) 'Returns "File" Object
	Msg = "File : " & FilePath & " Uses " & FileObj.Size & " Bytes" '.Size Property Returns Size Of The File In Bytes.
Else 'File Doesn't Exit.
	Msg = "File : " & FilePath & " Doesn't Exist."
End If
MsgBox Msg