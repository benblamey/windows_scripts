option explicit
dim WS, WN, fso, UN, dater, strDesktop, personal,favorites, CD_Burning, My_Music,My_Pictures,My_Video,strPrograms,strAUD,desktop,commonstartup,commondesktop,commonprograms,commonstartmenu
Set WS = WScript.CreateObject("WScript.Shell")
Set WN = WScript.CreateObject("WScript.Network")
Set fso = CreateObject("Scripting.FileSystemObject")
'******************************************************************************
Const bfnRoot = "D:\bfn_data\"
'******************************************************************************
strDesktop = WS.SpecialFolders("Desktop")
strPrograms = WS.SpecialFolders("Programs")
strAUD = WS.SpecialFolders("AllUsersDesktop")
UN=UCASE(WN.UserName)
'******************************************************************************
WS.RegWrite "HKEY_CURRENT_USER\Software\GNU\GnuPG\HomeDir","D:\bfn_data\gnupg","REG_SZ"
if UN="BEN" or UN="ADMIN" then
        'turn on num-lock for logon window
	WS.RegWrite "HKEY_USERS\.Default\Control Panel\Keyboard\InitialKeyboardIndicators","2","REG_SZ"  
	WS.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\ForceAutoLogon","1","REG_SZ"
	on error resume next
	' Remove Shared Docs from My Computer
	WS.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\DelegateFolders\{59031a47-3f72-44a7-89c5-5595fe6b30ee}\"
	on error goto 0

commondesktop=bfnroot & "AllUsers\Desktop"
commonstartup=bfnroot & "AllUsers\Start Menu\Programs\Startup"
commonprograms=bfnroot & "AllUsers\Start Menu\Programs"
commonstartmenu=bfnroot & "AllUsers\Start Menu"

WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Desktop",commondesktop,"REG_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Start Menu",commonstartmenu,"REG_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Startup",commonstartup,"REG_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Programs",commonprograms,"REG_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Desktop",commondesktop,"REG_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Start Menu",commonstartmenu,"REG_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Startup",commonstartup,"REG_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Programs",commonprograms,"REG_SZ"

end if        


' kill the search dog
WS.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState\Use Search Asst","no","REG_SZ"


'set shell folder paths
desktop = bfnRoot  & UN & "\desktop"
Personal = "D:\DOCS"
Favorites=bfnRoot & "FAVORITES"
My_Music="D:\docs\music"
My_Pictures="D:\docs\My Pictures"
My_Video="D:\docs\My Video"
' write user shell folder locations to registry
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal",Personal,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favorites",FAVORITES,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Music",My_Music,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures",My_Pictures,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Video",My_Video,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Desktop",Desktop,"REG_SZ"
'
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal",Personal,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites",FAVORITES,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Music",My_Music,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures",My_Pictures,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Video",My_Video,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Desktop",Desktop,"REG_SZ"
'******************************************************************************
' delete unneeded files
on error resume next
FSO.DeleteFile(strPrograms & "\Outlook Express.lnk")
'delete opera shortcuts on users desktop
QLaunch="C:\Documents and Settings\" & WN.UserName & "\Application Data\Microsoft\Internet Explorer\Quick Launch\"
FSO.DeleteFile(QLaunch & "QuickTime Player.lnk")
FSO.DeleteFile(QLaunch & "Launch Outlook Express.lnk")
FSO.DeleteFile(QLaunch & "Media Player Classic.lnk")
FSO.DeleteFile(strAUD & "\Quicktime Player.lnk")
on error goto 0
'###############################################################################
Wscript.Quit
Function RegRead(ByVal sRegValue)
On Error Resume Next
RegRead = WS.RegRead(sRegValue)
' If the value does not exist, error is raised
If Err Then
	RegRead = "01/01/1986"
	Err.clear
End If
' If a value is present but uninitialized the RegRead method
' returns the input value in Win2k.
If VarType(RegRead) < vbArray then
	If RegRead = sRegValue Then RegRead = ""
End If
On Error Goto 0
End Function
