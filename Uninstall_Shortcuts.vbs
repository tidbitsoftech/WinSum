Option Explicit

Dim fso, WshShell, strSendTo
Dim strMsg, strContinue, strTitle, strFile
Dim strSuccess, strFailed

strTitle = "Uninstall Shortcuts"

' Explain to the user what is about to happen...
strMsg = "You are about to delete the shortcuts in your SendTo folder" & vbCrLF
strMsg = strMsg & "for the Checksum scripts.  Do you wish to continue?" & vbCrLF & vbCrLF
strMsg = strMsg & "Click OK to continue or Cancel to quit."

strContinue =  MsgBox (strMsg,vbOKCancel,strTitle)
If strContinue <> 1 Then    'if user input equals anything other than OK, then quit.
	MsgBox "Operation has been cancelled."
	Wscript.Quit
End If

' Get the SendTo folder of the user.
set WshShell = Wscript.CreateObject("WScript.Shell")
strSendTo = WshShell.SpecialFolders("SendTo")
set WshShell = nothing

strSuccess = ""
strFailed = ""

set fso = CreateObject("Scripting.FileSystemObject")

delete_shortcut "md5sum.vbs"
delete_shortcut "sha1sum.vbs"
delete_shortcut "sha256sum.vbs"
delete_shortcut "sha512sum.vbs"

set fso = nothing

If strSuccess <> "" Then
	MsgBox "The following shortcuts were successfully removed:" & vbCrLF & strSuccess,,strTitle
End If

If strFailed <> "" Then
	MsgBox "The following shortcuts were not found:" & vbCrLF & strFailed,,strTitle
End If

Wscript.Quit


sub delete_shortcut (shortcut_name)
	strFile = strSendTo & "\" & shortcut_name & ".lnk"
	If fso.FileExists(strFile) Then		' verify the shortcut exists. If it doesn't then an error would be generated when we try to delete it.
		fso.DeleteFile(strFile)		' delete the shortcut
		strSuccess = strSuccess & vbCrLF & shortcut_name
	Else
		strFailed = strFailed & vbCrLF & shortcut_name
	End if
end sub
