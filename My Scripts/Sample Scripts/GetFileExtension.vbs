'This Example Demonstrates How To Use The "GetExtensionName" Method To 
'Get The File Extension Of The Last Component In A Specified Path.

Dim Fso, FileExtension
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FilePath = InputBox("Enter Full Path Of The File : ","File Path") 'Get Path Of The File
FileExtension = Fso.GetExtensionName(FilePath) 'Returns Extension Of The Specified File.
MsgBox "Extension Of : " & FilePath & " Is - " & FileExtension