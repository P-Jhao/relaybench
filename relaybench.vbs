Option Explicit

Dim shell, fso, projectRoot, cmd, lockPath, lockTimeoutMinutes, lockAgeMinutes
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

projectRoot = fso.GetParentFolderName(WScript.ScriptFullName)
lockPath = projectRoot & "\.relaybench-startup.lock"
lockTimeoutMinutes = 10

If fso.FileExists(lockPath) Then
  lockAgeMinutes = DateDiff("n", fso.GetFile(lockPath).DateLastModified, Now)
  If lockAgeMinutes < lockTimeoutMinutes Then
    MsgBox "RelayBench is starting or already running. Please wait.", vbInformation, "RelayBench"
    WScript.Quit 0
  End If

  ' Stale lock guard for abnormal exits.
  On Error Resume Next
  fso.DeleteFile lockPath, True
  On Error GoTo 0
End If

cmd = "cmd /c cd /d """ & projectRoot & """ && " & _
      "echo startup> """ & lockPath & """ && " & _
      "pnpm dev & del /q """ & lockPath & """ >nul 2>nul"

' 0 = hidden window, False = do not wait
shell.Run cmd, 0, False
