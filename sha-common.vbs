Option Explicit

' This file contains the common part used by all of the desired checksum scripts.
' I have separated this out because any change in functionality can be done in one file rather than having to modify multiple files.
'
' The script recognizes whether it is being run via WScript or Cscript and outputs accordingly.

Dim strFile, strText, strHash, strUserHash, strHashAlgorithm, strCompareHash, strMsg, strTitle
Dim is_cscript_running
Dim int_StartPos, intLength
Dim objShell, objExecObject

' Do you want to compare calculated hash with known hash
' Set this to TRUE if you want to be able to compare a known hash with the calculated hash.
' Set to FALSE if you just want the calculated has to be displayed.
strCompareHash = TRUE


' Check if running under "cscript".
' The advantage of using cscript is you get to copy the result.  WScript doesn't allow you to copy from the msgbox.
' This is assuming you are using a shortcut to drive the script.  This isn't needed if running straight from a DOS box (other than to correctly format that output).
is_cscript_running = FALSE
If InStr(1, WScript.FullName, "cscript.exe", vbTextCompare) <> 0 Then
	is_cscript_running = TRUE
End If

'If Wscript.Arguments.Count = 2 Then
	'Store the arguments.
	strFile = WScript.Arguments(0)
	'strHashAlgorithm= WScript.Arguments(1)

	If is_cscript_running Then
		Wscript.Echo vbCrLF & "Calculating hash..." & vbCrLF
	End If

	Set objShell = WScript.CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("cmd /c certutil -hashfile """ & strFile & """ " & strHashAlgorithm) ' "/c" causes the resulting window to close when it's finished.

	' Extract the file name from the full path
	int_StartPos = InstrRev(strFile, "\")		' Determine the starting position of the file name. First look for the path's trailing "\"
	If int_StartPos <> 0 Then		' If equal to 0, then there is no path info to remove and nothing to process
		int_StartPos = int_StartPos + 1		' Add 1 to the position
		intLength = (len(strFile) - int_StartPos) + 1		' Determine the length of the file name by subtracting the starting position from the total length of the line
		strFile = Mid( strFile, int_StartPos, intLength)
	End If

	' Extract the hash
	Do While Not objExecObject.StdOut.AtEndOfStream
		' certutil outputs 3 lines.  The actual hash is the only line that does NOT have the word "hash" in it.
		strText = objExecObject.StdOut.ReadLine()
		If Instr(strText, "hash") = 0 Then	' Get the hash
			strHash = strText
			Exit Do			' Now that we have the hash, exit the loop
		End If
	Loop

	set objShell = Nothing
	set objExecObject = Nothing
	
	strHash = replace(strHash, " ", "")		' Remove the spaces from the resulting hash.
	strHash = lcase(strHash)	' convert to lower case
	
	If strCompareHash Then
		strUserHash = trim(Inputbox(strHashAlgorithm & " hash of file: " & strFile & vbCrLF & strHash & vbCrLF & vbCrLF & "Enter known hash below to compare:"))
		strUserHash = lcase(strUserHash)	' convert to lower case
		
		if strUserHash <> "" Then
			If strUserHash = strHash Then
				strMsg = "The " & strHashAlgorithm & " checksums are the same." & vbCrLF
			Else
				strMsg = "The " & strHashAlgorithm & " checksums are -DIFFERENT-!" & vbCrLF
			End if
			strMsg = strMsg & vbCrLF & "File's hash: " & vbCrLF & strHash & vbCrLF & "You entered: " & vbCrLF & strUserHash
			WScript.Echo strMsg
		Else
			'If is_cscript_running Then
				' If running under cscript and comparing hashes, we still want to print the calculated hash
				' even if the user does not enter anything into the inputbox
				Wscript.Echo strHashAlgorithm & " hash of file: " & vbCrLF & strFile & vbCrLF & vbCrLF & strHash
			'End If
		End If
	Else
		' Output the file's hash
		WScript.Echo strHashAlgorithm & " hash of file: " & vbCrLF & strFile & vbCrLF & vbCrLF & strHash
	End If

	' If we are running from cscript, then don't close/quit until user has hada a chance to see results.
	If is_cscript_running Then
		WScript.Echo vbCrLF & vbCrLF & "Press ENTER to close..."
		WScript.StdIn.ReadLine()	' Check for the user to hit ENTER.
	End If

'Else
'  Wscript.Echo "hash.vbs usage: hash.vbs <source file> <MD5 | SHA1 | SHA256 | SHA384 | SHA512>"
'  Wscript.Echo "example: hash.vbs c:\test.txt SHA256"
'End If