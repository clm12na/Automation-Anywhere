''This Example Demonstrates How To Get Creation Date Of Any File.
'It Uses ".DateCreated" Property Of "File" Object.

'Also Demonstrates Use Of Following Methods Of "File" Object.
'".Copy (<<File Path>>)" -> To Copy File.
'".Move (<<File Path>>)" -> To Move File.
'".Delete" -> To Delete File.

Dim Fso, Msg, FileObj, FilePath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FilePath = InputBox("Enter Path Of The File : ","File Name") 'Get The Required File From The User.
If (Fso.FileExists(FilePath)) Then 'Checks Whether File Exits At The Specified Path
	Set FileObj = Fso.GetFile(FilePath) 'Returns "File" Object
	Msg = FilePath & " Was Created On " & FileObj.DateCreated

	If (FileObj.DateCreated < (Date - 2)) Then 'If File Is More Than 2 Days Old
		MsgBox FilePath & " Is More Than 2 Days Old."
		'----------------------------------------------------------------------
		'Uncomment The Code As Per Your Requirement.

		'Move File (Moves The File To Specified Destination).
		'Dim MovePath
		'MovePath = InputBox("Enter Path Where Files Are To Be Moved : ","File Path") 'Get The Path Where File Is To Be Moved
		'If Trim(MovePath) <> "" Then FileObj.Move MovePath 'Move Files To Specified Path

		'Copy File (Copies File To The Specified Destination).
		'Dim CopyPath
		'CopyPath = InputBox("Enter Path Where Files Are To Be Copied : ","File Path") 'Get The Path Where File Is To Be Copied
		'If Trim(CopyPath) <> "" Then FileObj.Copy CopyPath, True 'True -> Overwrite Existing File (If Present).

		'Delete File (Deletes File).
		'FileObj.Delete True 'True -> Force Deletion
		'----------------------------------------------------------------------
	End If
Else 'File Doesn't Exit.
	Msg = "File : " & FilePath & " Doesn't Exist."
End If
MsgBox Msg