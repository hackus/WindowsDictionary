' https://www.thewindowsclub.com/list-all-assigned-shortcut-keys-for-shortcuts-in-windows
Option Explicit

Dim objFSO, WshShell, arrFolders, fldr, strHotKey, outFilePath, outStream
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' Initialize hotkey string
strHotKey = ""

' Folders to scan
arrFolders = Array( _
    WshShell.SpecialFolders("AllUsersDesktop"), _
    WshShell.SpecialFolders("Desktop"), _
    WshShell.SpecialFolders("AllUsersStartMenu"), _
    WshShell.SpecialFolders("StartMenu"), _
    WshShell.SpecialFolders("AppData") & "\Microsoft\Internet Explorer\Quick Launch" _
)

' Optional: file to save results (in same folder as script)
outFilePath = objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), "ShortcutHotkeys.txt")
Set outStream = objFSO.CreateTextFile(outFilePath, True)

' Scan each folder
For Each fldr In arrFolders
    If objFSO.FolderExists(fldr) Then
        ScanFolder fldr
    End If
Next

outStream.Close

' Display results
If strHotKey = "" Then
    WshShell.Popup "No shortcut Hotkeys found.", 5, "Shortcut Hotkeys", 64
Else
    WshShell.Popup strHotKey, 20, "Shortcut Hotkeys Found", 64
End If

WScript.Echo "Results also saved to:" & vbCrLf & outFilePath

' Clean up
Set WshShell = Nothing
Set objFSO = Nothing

' === Recursive scanning subroutine ===
Sub ScanFolder(folderPath)
    Dim objFolder, colFiles, colSubFolders, objFile, objSubFolder, oShellLink, hotkey

    On Error Resume Next
    Set objFolder = objFSO.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub ' skip folder if access denied
    End If

    Set colFiles = objFolder.Files
    For Each objFile In colFiles
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "lnk" Then
            ' Try to read shortcut safely
            On Error Resume Next
            Set oShellLink = WshShell.CreateShortcut(objFile.Path)
            If Err.Number = 0 Then
                hotkey = Trim(oShellLink.Hotkey)
                If hotkey <> "" Then
                    strHotKey = strHotKey & "[" & hotkey & "]" & vbCrLf & objFile.Path & vbCrLf & vbCrLf
                End If
				outStream.WriteLine "[" & hotkey & "] " & objFile.Path
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next

    ' Recurse into subfolders
    Set colSubFolders = objFolder.SubFolders
    For Each objSubFolder In colSubFolders
        ScanFolder objSubFolder
    Next
End Sub
