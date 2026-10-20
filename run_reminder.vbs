' Launches reminder.py hidden (no console window). Works from whatever folder it lives in.
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
If fso.FileExists(dir & "\disabled.flag") Then WScript.Quit
CreateObject("WScript.Shell").Run "pythonw """ & dir & "\reminder.py""", 0, False
