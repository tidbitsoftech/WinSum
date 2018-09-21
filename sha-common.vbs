Option Explicit

' This file contains the common part used by all of the desired checksum scripts.
' I have separated this out because any change in functionality can be done in one file rather than having to modify multiple files.

Dim strFile, strText, strHash, strUserHash, strHashAlgorithm, strCompareHash, strMsg, strTitle, strVersion
Dim int_StartPos, intLength
Dim objShell, objExecObject

strVersion = "1.5"

' Do you want to compare calculated hash with known hash
' Set this to TRUE if you want to be able to compare a known hash with the calculated hash.
' Set to FALSE if you just want the calculated has to be displayed.
strCompareHash = TRUE

' Set the title used in the MsgBox windows
strTitle = strHashAlgorithm & " checksum v" & strVersion

' Check that a file has been specified.
If Wscript.Arguments.Count = 0 Then
	MsgBox "A file was not supplied.",,strTitle
	Wscript.Quit
End if

'Store the arguments.
strFile = WScript.Arguments(0)
'strHashAlgorithm= WScript.Arguments(1)

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

' Extract the file name from the full path
int_StartPos = InstrRev(strFile, "\")		' Determine the starting position of the file name. First look for the path's trailing "\"
If int_StartPos <> 0 Then		' If equal to 0, then there is no path info to remove and nothing to process
	int_StartPos = int_StartPos		' Add 1 to the position
	intLength = (len(strFile) - int_StartPos)		' Determine the length of the file name by subtracting the starting position from the total length of the line
	strFile = Mid( strFile, (int_StartPos + 1), intLength)
End If

strHash = replace(strHash, " ", "")		' Remove the spaces from the resulting hash.
strHash = lcase(strHash)	' convert to lower case

If strCompareHash Then
	strUserHash = trim(Inputbox(strHashAlgorithm & " hash of file:" & vbCrLF & strFile & vbCrLF & vbCrLF & strHash & vbCrLF & "Enter known hash below to compare:", strTitle))
	strUserHash = lcase(strUserHash)	' convert the hash entered by the user to lower case
	
	if strUserHash <> "" Then
		If strUserHash = strHash Then
			strMsg = "The " & strHashAlgorithm & " checksums are the same." & vbCrLF
		Else
			strMsg = "The " & strHashAlgorithm & " checksums are -DIFFERENT-!" & vbCrLF
		End if
		strMsg = strMsg & vbCrLF & "File's hash: " & vbCrLF & strHash & vbCrLF & "You entered: " & vbCrLF & strUserHash
		MsgBox strMsg,,strTitle
	Else
		'Dim strDontCare		'Used as a throw-away variable
		'strMsg = "Use CTRL-C to copy the resulting checksum..." & vbCrLF & vbCrLF & vbCrLF
		'strMsg = strMsg & strHashAlgorithm & " checksum of file:" & vbCrLF
		'strMsg = strMsg & strFile
		'strDontCare = Inputbox(strMsg, strTitle, strHash )
		'================
		MsgBox strFile & vbCrLF & vbCrLF & strHash,,strTitle
	End If
Else
	' Output the file's hash
	MsgBox strFile & vbCrLF & vbCrLF & strHash,,strTitle
End If
