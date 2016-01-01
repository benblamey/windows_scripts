Set WS = WScript.CreateObject("WScript.Shell")
Set WN = WScript.CreateObject("WScript.Network")

const WinLogRoot = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\"

strPassAdmin = InputBox("Enter Password for Administrator, Cancel to Cancel", "Set Password", "password")
if strPassAdmin = "" then Wscript.Quit
Wscript.Sleep(500)

strPassAdministrator = InputBox("Enter Password for Admin, Cancel to Cancel", "Set Password", "password")
if strPassAdministrator = "" then Wscript.Quit
Wscript.Sleep(500)

strPassUser = InputBox("Enter Password for User, Cancel to Cancel", "Set Password", "password")
if strPassUser = "" then Wscript.Quit

WS.RegWrite "HKLM\Software\BFN\PassAdmin",strPassAdmin,"REG_SZ"
WS.RegWrite "HKLM\Software\BFN\PassUser",strPassUser,"REG_SZ"

'for old compatability
WS.RegWrite "HKLM\Software\BFN\strPassB",strPassUser,"REG_SZ"
WS.RegWrite "HKLM\Software\BFN\strPassA",strPassAdmin,"REG_SZ"

WS.Run "cmd /k net user administrator " & strPassAdministrator, ,-1
WS.Run "cmd /k net user admin " & strPassAdmin, , -1
WS.Run "cmd /k net user user " & strPassUser, , -1

temp = MsgBox("Passwords have been set on computer: "&WN.ComputerName&" , you may wish to print this: "& Chr(13) & Chr(10)& Chr(13) & Chr(10) & "Administrator: " & strPAssAdministrator & Chr(13) & Chr(10) & "Admin: " & strPAssAdmin & Chr(13) & Chr(10) & "User: " & strPassUser & Chr(13) & Chr(10) & Chr(13) & Chr(10)& Chr(13) & Chr(10)  & "(Hit Ctrl-C Now,Run Notepad,Control-V,File,Print)")


UN=UCASE(WN.UserName)
const AUTOLOGONSTATUS = 1
if UN="ADMIN" then
	WS.RegWrite WinLogRoot&"DefaultUserName","admin","REG_SZ"
	WS.RegWrite WinLogRoot&"AutoAdminLogon",AUTOLOGONSTATUS,"REG_SZ"
	WS.RegWrite WinLogRoot&"DefaultPassword",strPassAdmin,"REG_SZ"
end if
if UN="USER" then
	WS.RegWrite WinLogRoot&"DefaultUserName","USER","REG_SZ"
	WS.RegWrite WinLogRoot&"AutoAdminLogon",AUTOLOGONSTATUS,"REG_SZ"
	WS.RegWrite WinLogRoot&"DefaultPassword",strPassUser,"REG_SZ"
end if
if UN="ADMINISTRATOR" then
	WS.RegWrite WinLogRoot&"DefaultUserName","ADMINISTRATOR","REG_SZ"
	WS.RegWrite WinLogRoot&"AutoAdminLogon",AUTOLOGONSTATUS,"REG_SZ"
	WS.RegWrite WinLogRoot&"DefaultPassword",strPassAdministrator,"REG_SZ"
end if
