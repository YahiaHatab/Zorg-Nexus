Set WshShell = CreateObject("WScript.Shell")
Set WMI = GetObject("winmgmts:\\.\root\cimv2")

' Function to check if a specific node script is running
Function IsScriptRunning(scriptName)
    Dim colProcesses, objProcess
    Set colProcesses = WMI.ExecQuery("Select * from Win32_Process Where Name = 'node.exe'")
    IsScriptRunning = False
    For Each objProcess In colProcesses
        If InStr(1, objProcess.CommandLine, scriptName, vbTextCompare) > 0 Then
            IsScriptRunning = True
            Exit Function
        End If
    Next
End Function

' Only run if not already active
If Not IsScriptRunning("Zorg-Nexus\server.js") Then
    WshShell.Run "node ""C:\Users\Yehya\Desktop\Zorg-Nexus\server.js""", 0, False
End If

If Not IsScriptRunning("ZORG-main\app.js") Then
    WshShell.Run "node ""C:\Users\Yehya\Desktop\ZORG-main\app.js""", 0, False
End If

Set WshShell = Nothing
Set WMI = Nothing