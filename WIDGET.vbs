Set WshShell = CreateObject("WScript.Shell")
Set WMI = GetObject("winmgmts:\\.\root\cimv2")

Function IsProcessRunning(procName, cmdSubstr)
    Dim colProcesses, objProcess
    Set colProcesses = WMI.ExecQuery("Select * from Win32_Process Where Name = '" & procName & "'")
    IsProcessRunning = False
    For Each objProcess In colProcesses
        If InStr(1, objProcess.CommandLine, cmdSubstr, vbTextCompare) > 0 Then
            IsProcessRunning = True
            Exit Function
        End If
    Next
End Function

' 1. Ensure server.js is running
If Not IsProcessRunning("node.exe", "server.js") Then
    WshShell.Run "cmd.exe /c node ""C:\Users\Yehya\Desktop\Zorg-Nexus\server.js""", 0, False
End If

' 2. Launch Electron Desktop Widget
If Not IsProcessRunning("electron.exe", "widget/main.js") Then
    WshShell.Run "cmd.exe /c npx electron ""C:\Users\Yehya\Desktop\Zorg-Nexus\widget\main.js""", 0, False
End If

Set WshShell = Nothing
Set WMI = Nothing
