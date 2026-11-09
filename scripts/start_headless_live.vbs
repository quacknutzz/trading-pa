Set FSO = CreateObject("Scripting.FileSystemObject")
ScriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)
RootDir = FSO.GetParentFolderName(ScriptDir)

Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = RootDir
WshShell.Run "cmd /c """ & ScriptDir & "\run_background_live.bat""", 0, False
Set WshShell = Nothing
