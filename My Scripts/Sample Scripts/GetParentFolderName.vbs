'This Example Demonstrates How To Use The "GetParentFolderName" Method 
'To Get The Name Of The Parent Folder Of A Specified Path.

Dim Fso, FilePath, ParentFolderPath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FilePath = InputBox("Enter Full Path Of The File : ","File Path") 'Get Path Of The File
ParentFolderPath = Fso.GetParentFolderName(FilePath) 'Returns Parent Folder Name
MsgBox "Parent Folder : " & ParentFolderPath