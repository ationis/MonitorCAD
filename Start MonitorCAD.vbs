Option Explicit

Dim shell, fileSystem, projectFolder, pythonwPath, scriptPath, command
Set shell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")

projectFolder = fileSystem.GetParentFolderName(WScript.ScriptFullName)
pythonwPath = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & _
    "\Programs\Python\Python313\pythonw.exe"
scriptPath = fileSystem.BuildPath(projectFolder, "monitorcad_v3.py")

If Not fileSystem.FileExists(pythonwPath) Then
    MsgBox "Python 3.13 was not found:" & vbCrLf & pythonwPath, _
        vbCritical, "MonitorCAD v3"
    WScript.Quit 1
End If

If Not fileSystem.FileExists(scriptPath) Then
    MsgBox "MonitorCAD script was not found:" & vbCrLf & scriptPath, _
        vbCritical, "MonitorCAD v3"
    WScript.Quit 1
End If

command = """" & pythonwPath & """ """ & scriptPath & """"
shell.CurrentDirectory = projectFolder
shell.Run command, 0, False
