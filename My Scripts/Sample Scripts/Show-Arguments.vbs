'This Example Demonstrates How To Get arguement passed to the VBScript.

If WScript.Arguments.Count > 0 Then
	MsgBox WScript.Arguments.Item(0)
Else
	MsgBox "Please pass a parameter to this script"
End if