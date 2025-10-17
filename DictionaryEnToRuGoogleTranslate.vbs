Option Explicit
Include "Common"

Dim iURL 
Dim objShell
Dim result

' === Create a shortcut to this script and let user assign a hotkey ===
Call CreateSelfShortcut()

result = ClipBoard(Null)
iURL = "https://translate.google.com/#en/ru/" & """" & result & """"
Set objShell = WScript.CreateObject( "Shell.Application" )
' objShell.ShellExecute "C:\Program Files\Mozilla Firefox\firefox.exe", iURL, "", "", 1
objShell.ShellExecute "C:\Program Files\Google\Chrome\Application\chrome.exe", iURL, "", "", 1

Set objShell = Nothing

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