Option Explicit

Dim strHashAlgorithm, strPath
Dim fso, CurrentDirectory

' Set the hashing algorithm
' the available hashes are <MD5 | SHA1 | SHA256 | SHA384 | SHA512>.  See "certutil" documentation for further details.
strHashAlgorithm = "SHA1"   ' NOTE: this must be in ALL CAPS.  If not, certutil will produce an error.

strPath = Wscript.ScriptFullName		' Determine the script's path and filename so we can include the common file later.
set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetParentFolderName(strPath)		' Determine just the path information
set fso = nothing

' now that we know the path to the script, we can call the common file to be included
' as it should be stored in the same directory as this script.
includeFile CurrentDirectory & "\sha-common.vbs"

sub includeFile (fSpec)
    dim fileSys, file, fileData
    set fileSys = createObject ("Scripting.FileSystemObject")
    set file = fileSys.openTextFile (fSpec)
    fileData = file.readAll ()
    file.close
    executeGlobal fileData
    set file = nothing
    set fileSys = nothing
end sub