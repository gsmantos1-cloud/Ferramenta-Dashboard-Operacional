Dim fso, dir, bat
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
bat = dir & "\gs_bg.bat"
CreateObject("WScript.Shell").Run "cmd /c """ & bat & """", 0, False
