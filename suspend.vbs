Set objShell = CreateObject("WScript.Shell")

' Poner el PC en modo de suspensión
objShell.Run "rundll32.exe powrprof.dll,SetSuspendState 0,1,0", 0, True