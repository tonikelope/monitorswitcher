Set WshShell = CreateObject("WScript.Shell")
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = "powershell.exe -ExecutionPolicy Bypass -NoProfile -File """ & scriptPath & "\MonitorSwitcher.ps1"""
WshShell.Run command, 0, False
