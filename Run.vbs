Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd.exe /c node ""C:\Users\Yehya\Desktop\Zorg-Nexus\server.js""", 0, False
Set WshShell = Nothing