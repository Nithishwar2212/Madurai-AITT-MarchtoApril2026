Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonwPath = "C:\Program Files\Python311\pythonw.exe"
appPath = fso.BuildPath(scriptDir, "gstr1_tool.py")

shell.Run """" & pythonwPath & """ """ & appPath & """", 0, False
