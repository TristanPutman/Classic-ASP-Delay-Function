Sub Delay(seconds)
	set oShell = CreateObject("Wscript.Shell")
	cmd = "%COMSPEC% /c ping -n " & 1 + seconds & " 127.0.0.1>nul"
	oShell.Run cmd,0,1
End Sub
