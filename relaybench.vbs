Option Explicit

Dim shell, fso, projectRoot, cmd
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

projectRoot = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "cmd /c cd /d """ & projectRoot & """ && pnpm dev"

' 0 = hidden window, False = do not wait
shell.Run cmd, 0, False

