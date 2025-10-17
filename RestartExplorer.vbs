Option Explicit

' Example main script
Call RestartExplorerSafe()

' WScript.Echo "Explorer restarted successfully."


Sub RestartExplorerSafe()
    Dim shell, exec, retries
    Set shell = CreateObject("WScript.Shell")

    On Error Resume Next
    shell.Run "taskkill /f /im explorer.exe", 0, True
    On Error GoTo 0

    ' Wait until it's really stopped
    retries = 0
    Do While IsProcessRunning("explorer.exe") And retries < 10
        WScript.Sleep 500
        retries = retries + 1
    Loop

    ' Restart
    shell.Run "explorer.exe", 0, False
End Sub

Sub RestartExplorerAndOpenFolders()
    Dim shell, fso, userFolder, startMenuPath, scriptFolder

    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    userFolder = shell.ExpandEnvironmentStrings("%AppData%")
    startMenuPath = fso.BuildPath(userFolder, "Microsoft\Windows\Start Menu\Programs")
    scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)

    ' --- Kill explorer.exe ---
    On Error Resume Next
    shell.Run "taskkill /f /im explorer.exe", 0, True
    On Error GoTo 0

    WScript.Sleep 1500  ' give time to fully exit

    ' --- Restart explorer.exe ---
    shell.Run "explorer.exe", 0, False
    WScript.Sleep 2000  ' allow Explorer to start up properly

    ' --- Open desired folders ---
    shell.Run "explorer.exe """ & userFolder & """", 1, False
    shell.Run "explorer.exe """ & startMenuPath & """", 1, False
    shell.Run "explorer.exe """ & scriptFolder & """", 1, False
End Sub

Function IsProcessRunning(procName)
    Dim shell, exec, output
    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec("cmd /c tasklist /fi ""imagename eq " & procName & """ | find /i """ & procName & """")
    output = LCase(exec.StdOut.ReadAll)
    IsProcessRunning = InStr(output, LCase(procName)) > 0
End Function
