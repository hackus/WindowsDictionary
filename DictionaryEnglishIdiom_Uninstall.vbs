Option Explicit
Include "Common"

Dim iURL 
Dim objShell
Dim result
Dim JSEngine,ws,WSC, sEncoded

' === Create a shortcut to this script and let user assign a hotkey ===
Call RemoveSelfShortcut()

' This is the Sub that opens external files and reads in the contents.
' In this way, you can have separate files for data and libraries of functions
Sub Include(yourFile)
  Dim oFSO, oFileBeingReadIn   ' define Objects
  Dim sFileContents            ' define Strings
 
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFileBeingReadIn = oFSO.OpenTextFile(yourFile & ".vbs", 1)
  sFileContents = oFileBeingReadIn.ReadAll
  oFileBeingReadIn.Close
  ExecuteGlobal sFileContents
End Sub
