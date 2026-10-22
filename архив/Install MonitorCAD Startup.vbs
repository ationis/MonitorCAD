Option Explicit

Dim shell, fileSystem, projectFolder, launcherPath, startupFolder
Dim shortcutPath, shortcut
Set shell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")

projectFolder = fileSystem.GetParentFolderName(WScript.ScriptFullName)
launcherPath = fileSystem.BuildPath(projectFolder, "Start MonitorCAD.vbs")

If Not fileSystem.FileExists(launcherPath) Then
    MsgBox "MonitorCAD launcher was not found:" & vbCrLf & launcherPath, _
        vbCritical, "MonitorCAD startup"
    WScript.Quit 1
End If

startupFolder = shell.SpecialFolders("Startup")
shortcutPath = fileSystem.BuildPath(startupFolder, "MonitorCAD.lnk")
Set shortcut = shell.CreateShortcut(shortcutPath)
shortcut.TargetPath = launcherPath
shortcut.WorkingDirectory = projectFolder
shortcut.Description = "Start MonitorCAD when you sign in"
shortcut.Save

MsgBox "MonitorCAD was added to Windows startup." & vbCrLf & _
    "The app will wait for KOMPAS-3D if it is not open yet.", _
    vbInformation, "MonitorCAD startup"
