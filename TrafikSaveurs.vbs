' Lance Trafik Saveurs sans fenêtre noire (console) — double-cliquez !
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dir
sh.Run """" & dir & "\TrafikSaveurs.exe""", 1, False
