Option Explicit

Dim strPath, fso, strCurrentDirectory
Dim WshShell, strSendTo
Dim strMsg, strContinue, strTitle, strProg

strProg = "WinSum"
strTitle = strProg & " - Create Shortcuts"

' Explain what is about to happen...
strMsg = "You are about to create the " & strProg & " shortcuts in your" & vbCrLF
strMsg = strMsg & "SendTo folder.  If these shortcuts already exist," & vbCrLF
strMsg = strMsg & "they will be overwritten.  Do you wish to continue?" & vbCrLF & vbCrLF
strMsg = strMsg & "Click OK to continue or Cancel to quit."

strContinue =  MsgBox (strMsg,vbOKCancel,strTitle)
If strContinue <> 1 Then    'if user input equals anything other than OK, then quit.
	MsgBox "Operation has been cancelled.",,strTitle
	Wscript.Quit
End If

strPath = Wscript.ScriptFullName		' Determine the script's path and file name
set fso = CreateObject("Scripting.FileSystemObject")
strCurrentDirectory = fso.GetParentFolderName(strPath)		' Determine just the path information
set fso = nothing

set WshShell = Wscript.CreateObject("WScript.Shell")
strSendTo = WshShell.SpecialFolders("SendTo")

create_shortcut strProg & " - MD5sum",strProg & ".vbs","MD5"
create_shortcut strProg & " - sha1sum",strProg & ".vbs","SHA1"
create_shortcut strProg & " - sha256sum",strProg & ".vbs","SHA256"
create_shortcut strProg & " - sha512sum",strProg & ".vbs","SHA512"

set WshShell = nothing

MsgBox "The shortcuts have been created/updated.",,strTitle
Wscript.Quit

sub create_shortcut (strShortCutName, file_name, strAlg)
	Dim oMyShortCut
	set oMyShortCut= WshShell.CreateShortcut(strSendTo & "\" & strShortCutName & ".lnk")
	oMyShortCut.TargetPath = "WScript.exe"
	oMyShortCut.WorkingDirectory = strCurrentDirectory
	oMyShortCut.Arguments = """" & strCurrentDirectory & "\" & file_name & """ """ & strAlg & """"
	oMyShortCut.Save
	set oMyShortCut = nothing
end sub
