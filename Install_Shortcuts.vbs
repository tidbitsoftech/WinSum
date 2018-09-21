Option Explicit

Dim strPath, fso, strCurrentDirectory
Dim WshShell, strSendTo
Dim strMsg, strContinue, strTitle

strTitle = "Install Shortcuts"

' Explain to the user what is about to happen...
strMsg = "You are about to create shortcuts in your SendTo folder" & vbCrLF
strMsg = strMsg & "for the Checksum scripts.  If these shortcuts already exist," & vbCrLF
strMsg = strMsg & "they will be overwritten.  Do you wish to continue?" & vbCrLF & vbCrLF
strMsg = strMsg & "Click OK to continue or Cancel to quit."

strContinue =  MsgBox (strMsg,vbOKCancel,strTitle)
If strContinue <> 1 Then    'if user input equals anything other than OK, then quit.
    MsgBox "Operation has been cancelled."
    Wscript.Quit
End If

strPath = Wscript.ScriptFullName		' Determine the script's path and file name
set fso = CreateObject("Scripting.FileSystemObject")
strCurrentDirectory = fso.GetParentFolderName(strPath)		' Determine just the path information
set fso = nothing

set WshShell = Wscript.CreateObject("WScript.Shell")
strSendTo = WshShell.SpecialFolders("SendTo")
set WshShell = nothing

create_shortcut "md5sum.vbs"
create_shortcut "sha1sum.vbs"
create_shortcut "sha256sum.vbs"
create_shortcut "sha512sum.vbs"

MsgBox "The shortcuts have been created/updated.",,strTitle
Wscript.Quit

sub create_shortcut (file_name)
    Dim oMyShortCut
    set oMyShortCut= WshShell.CreateShortcut(strSendTo & "\" & file_name & ".lnk")
    oMyShortCut.TargetPath = strCurrentDirectory & "\" & file_name 
    oMyShortCut.WorkingDirectory = strCurrentDirectory
    oMyShortCut.Save
    set oMyShortCut = nothing
end sub
