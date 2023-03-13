Dim WinscriptHost
Set WinscriptHost = CreateObject("Wscript.Shell")
WinscriptHost.Run Chr(34) & "D:\Automation\StartSlave.bat" & Chr(34), 0
Set WinscriptHost = Nothing