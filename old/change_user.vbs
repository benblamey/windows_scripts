option explicit
Dim WS, WN,bfnRoot,strPassA,strPassB
Set WS = WScript.CreateObject("WScript.Shell")
Set WN = WScript.CreateObject("WScript.Network")

if UCASE(WN.UserName)="USER" then
	strPassB = WS.RegRead("HKLM\Software\BFN\PassAdmin")
	WS.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon","1","REG_SZ"
	WS.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName","admin","REG_SZ"
	WS.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword",strPassB,"REG_SZ"
else
	strPassA = WS.RegRead("HKLM\Software\BFN\PassUser")
	WS.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon","1","REG_SZ"
	WS.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword",strPassA,"REG_SZ"
	WS.RegWrite "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName","user","REG_SZ"
end if
Wscript.Sleep(1100)
WS.Run("logoff")
