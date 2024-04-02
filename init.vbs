Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ruta del archivo log.txt
strLogPath = "log.txt"

' Nombre del proceso a verificar
strProcessName = "cyberadmin.exe"

' Verificar si el proceso está corriendo
If CheckProcessRunning(strProcessName) Then
    WriteLog "El proceso " & strProcessName & " ya estaba abierto."
Else
		objShell.Run strProcessName, 0, True
    WriteLog "El proceso " & strProcessName & " se abrió con el script."
End If

' Esperar 25 segundos
WScript.Sleep 15000

' Lanzar acciones en CMD
strCmd1 = "cd ""c:\program files\shadow defender\"""
strCmd2 = "c:"
strCmd3 = "CmdTool.exe /enter:C /now"
strCmd4 = "CmdTool.exe /exit:C"

' Ejecutar comandos en CMD
ExecuteCMD strCmd1
ExecuteCMD strCmd2
ExecuteCMD strCmd3
ExecuteCMD strCmd4

' Función para verificar si un proceso está corriendo
Function CheckProcessRunning(strProcessName)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcessName & "'")

    If colProcesses.Count > 0 Then
        CheckProcessRunning = True
    Else
        CheckProcessRunning = False
    End If
End Function

' Función para escribir en el archivo log.txt
Sub WriteLog(strMessage)
    Set objLogFile = objFSO.OpenTextFile(strLogPath, 8, True)
    objLogFile.WriteLine Now & " - " & strMessage
    objLogFile.Close
End Sub

' Función para ejecutar comandos en CMD
Sub ExecuteCMD(strCommand)
    objShell.Run "cmd /c " & strCommand, 0, True
End Sub
