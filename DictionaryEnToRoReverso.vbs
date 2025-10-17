Option Explicit
Include "Common"

Dim iURL 
Dim objShell
Dim result
Dim JSEngine,ws,WSC, sEncoded

' === Create a shortcut to this script and let user assign a hotkey ===
Call CreateSelfShortcut()

Set ws = CreateObject("WScript.Shell")
WSC = ws.ExpandEnvironmentStrings("%AppData%\urlencdec.wsc")
Call Create_URL_ENC_DEC_Component(WSC)
Set JSEngine = GetObject("Script:"& WSC)

result = ClipBoard(Null)

' Show the URL encoded string:
sEncoded = JSEngine.encode(result)

iURL = "https://context.reverso.net/translation/romanian-english/" & sEncoded
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
