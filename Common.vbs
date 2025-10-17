Option Explicit

Function ClipBoard(input)
'@description: A quick way to set and get your clipboard.
'@author: Jeremy England (SimplyCoded)
  If IsNull(input) Then
    ClipBoard = CreateObject("HTMLFile").parentWindow.clipboardData.getData("Text")
    If IsNull(ClipBoard) Then ClipBoard = ""
  Else
    CreateObject("WScript.Shell").Run _
      "mshta.exe javascript:eval(""document.parentWindow.clipboardData.setData('text','" _
      & Replace(Replace(Replace(input, "'", "\\u0027"), """","\\u0022"),Chr(13),"\\r\\n") & "');window.close()"")", _
      0,True
  End If
End Function

Sub Create_URL_ENC_DEC_Component(WSC)
	Dim fso,File
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set File = fso.OpenTextFile(WSC,2,True)

	File.WriteLine "<?xml version=""1.0""?>"
	File.WriteLine "<component>"
	File.WriteLine "<?component error=""true"" debug=""true""?>"
	File.WriteLine     "<registration"
	File.WriteLine         "description=""Url Encode / Decode Helper"""
	File.WriteLine         "progid=""JSEngine.Url"""
	File.WriteLine         "version=""1.0"""
	File.WriteLine         "classid=""{80246bcc-45d4-4e92-95dc-4fd9a93d8529}"""
	File.WriteLine     "/>"
	File.WriteLine    "<public>"
	File.WriteLine         "<method name=""encode"">"
	File.WriteLine             "<PARAMETER name=""s""/>"
	File.WriteLine         "</method>"
	File.WriteLine         "<method name=""decode"">"
	File.WriteLine             "<PARAMETER name=""s""/>"
	File.WriteLine         "</method>"
	File.WriteLine     "</public>"
	File.WriteLine     "<script language=""JScript"">"
	File.WriteLine     "<![CDATA["
	File.WriteLine         "var description = new UrlEncodeDecodeHelper;"
	File.WriteLine         "function UrlEncodeDecodeHelper() {"
	File.WriteLine             "this.encode = encode;"
	File.WriteLine             "this.decode = decode;"
	File.WriteLine         "}"
	File.WriteLine         "function encode(s) {"
	File.WriteLine            "return encodeURIComponent(s).replace(/'/g,""%27"").replace(/""/g,""%22"");"
	File.WriteLine         "}"
	File.WriteLine         "function decode(s) {"
	File.WriteLine             "return decodeURIComponent(s.replace(/\+/g,  "" ""));"
	File.WriteLine         "}"
	File.WriteLine     "]]>"
	File.WriteLine     "</script>"
	File.WriteLine "</component>"

End Sub

Sub CreateSelfShortcut()
    Dim fso, wsh, shell
    Dim scriptPath, scriptName
    Dim shortcutPath, tempFolder, startMenuPath
    Dim shortcut, folder, item, dest, destFileName

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    Set shell = CreateObject("Shell.Application")

    ' Current script info
    scriptPath = WScript.ScriptFullName
    scriptName = fso.GetBaseName(scriptPath)
    tempFolder = fso.GetParentFolderName(scriptPath)

    ' Step 1: create shortcut in the same folder as the script
    shortcutPath = fso.BuildPath(tempFolder, scriptName & ".lnk")
	startMenuPath = wsh.SpecialFolders("Programs")
	dest = fso.BuildPath(startMenuPath, fso.GetFileName(shortcutPath))
	destFileName = fso.BuildPath(dest, scriptName & ".lnk")
	
	' WScript.Echo dest

    If Not fso.FileExists(dest) Then
	
		If fso.FileExists(fso.BuildPath(tempFolder, fso.GetFileName(shortcutPath))) Then
			fso.DeleteFile fso.BuildPath(tempFolder, fso.GetFileName(shortcutPath))
		End If
		
		Set shortcut = wsh.CreateShortcut(shortcutPath)
		shortcut.TargetPath = "wscript.exe"
		shortcut.Arguments = """" & scriptPath & """"
		shortcut.WorkingDirectory = tempFolder
		' shortcut.IconLocation = "shell32.dll,1"
		shortcut.IconLocation = wsh.ExpandEnvironmentStrings("%SystemRoot%\System32\WScript.exe,1")
		shortcut.Description = "Shortcut for " & scriptName
		shortcut.Save
		
		WScript.Sleep 200 ' small delay to ensure file is registered
		
		Set folder = shell.Namespace(tempFolder)
		Set item = folder.ParseName(fso.GetFileName(shortcutPath))
		
		' WScript.Echo "Only open Properties if no hotkey assigned"
		If Trim(shortcut.Hotkey) = "" Then
			WScript.Echo "Please assign a Hotkey for this shortcut."
			Set shell = CreateObject("Shell.Application")
			shell.ShellExecute shortcutPath, "", "", "properties", 1
		End If
		
		WScript.Sleep 200
		
		WScript.Echo "Step 3: move shortcut to Start Menu Programs folder"
		
		fso.MoveFile shortcutPath, dest
    End If
End Sub





