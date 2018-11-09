Option Explicit

Dim strFile, strText, strHash, strUserHash, strHashAlgorithm, strCompareHash, strMsg, strTitle, strVersion
Dim int_StartPos, intLength
Dim objShell, objExecObject

' Set the title and version used in the MsgBox windows
strVersion = "2.0"
strTitle = "WinSum v" & strVersion

' Set to TRUE if you want to be able to compare a known checksum with the calculated checksum (default).
' Set to FALSE if you just want the calculated checksum to be displayed.
strCompareHash = TRUE



If Wscript.Arguments.Count <= 1 Then		' Check that a file has been specified.
	MsgBox "Invalid number of arguments or a file was not supplied.",,strTitle
	Wscript.Quit
ElseIf Wscript.Arguments.Count > 2 Then		' Check that only one file is being passed.
	MsgBox "It looks like you tried to check more than one file.",,strTitle
	Wscript.Quit
End if

' Parse the arguments.
strHashAlgorithm = UCase(WScript.Arguments(0))	' The Algorithm has to be passed as uppercase or CertUtil will throw an error.
Select Case strHashAlgorithm
	Case "MD5","SHA1","SHA256","SHA512"		' These are valid
	Case Else								' Anything else is an error
		MsgBox "A valid algorithm was not specified.",,strTitle
		Wscript.Quit
End Select

strFile = WScript.Arguments(1)

Set objShell = WScript.CreateObject("WScript.Shell")
Set objExecObject = objShell.Exec("cmd /c certutil -hashfile """ & strFile & """ " & strHashAlgorithm) ' "/c" causes the resulting window to close when it's finished.
' Extract the hash from certutil's output
Do While Not objExecObject.StdOut.AtEndOfStream
	' certutil outputs 3 lines.  The actual hash is the only line that does NOT have the word "hash" in it.
	strText = objExecObject.StdOut.ReadLine()
	If Instr(strText, "hash") = 0 Then	' Get the hash
		strHash = strText
		Exit Do			' Now that we have the hash, exit the loop
	End If
Loop
set objExecObject = Nothing
set objShell = Nothing

strHash = replace(strHash, " ", "")		' Remove any spaces from the resulting hash.
strHash = lcase(strHash)	' convert to lower case

' Extract the file name from the full path
int_StartPos = InstrRev(strFile, "\")		' Determine the starting position of the file name. First look for the path's trailing "\"
If int_StartPos <> 0 Then		' If equal to 0, then there is no path info to remove and nothing to process
	int_StartPos = int_StartPos		' Add 1 to the position
	intLength = (len(strFile) - int_StartPos)		' Determine the length of the file name by subtracting the starting position from the total length of the line
	strFile = Mid( strFile, (int_StartPos + 1), intLength)
End If

If strCompareHash Then
	strUserHash = trim(Inputbox(strHashAlgorithm & " checksum of file:" & vbCrLF & strFile & vbCrLF & vbCrLF & strHash & vbCrLF & "Enter known checksum below to compare:", strTitle))
	strUserHash = lcase(strUserHash)	' convert the checksum entered by the user to lower case

	if strUserHash <> "" Then
		If strUserHash = strHash Then
			strMsg = "The " & strHashAlgorithm & " checksums are the same." & vbCrLF
		Else
			strMsg = "The " & strHashAlgorithm & " checksums are -DIFFERENT-!" & vbCrLF
		End if
		strMsg = strMsg & vbCrLF & "File's checksum: " & vbCrLF & strHash & vbCrLF & "You entered: " & vbCrLF & strUserHash
		MsgBox strMsg,,strTitle
	Else
		MsgBox strFile & vbCrLF & vbCrLF & strHash,,strTitle
	End If
Else
	' Output the file's hash
	MsgBox strFile & vbCrLF & vbCrLF & strHash,,strTitle
End If
