Option Explicit

Dim fso, folder, file, scriptFolder, shell
Dim countRemoved

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Get current folder path (where this script resides)
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Set folder = fso.GetFolder(scriptFolder)

countRemoved = 0

WScript.Echo "Searching for *Uninstall.vbs scripts in: " & scriptFolder

' Loop through all files in the current folder
For Each file In folder.Files
    If LCase(Right(file.Name, 14)) = "_uninstall.vbs" Then
        ' WScript.Echo "Running: " & file.Name
        shell.Run "wscript.exe """ & file.Path & """", 1, True
        countRemoved = countRemoved + 1
    End If
Next

If countRemoved = 0 Then
    WScript.Echo "No *Uninstall.vbs scripts found."
Else
    WScript.Echo vbCrLf & "Done! " & countRemoved & " shortcut(s) removed."
End If
