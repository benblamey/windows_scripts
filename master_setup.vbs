dim loc
loc=SetLocale("en-gb")
'************************************
' (c) Ben Blamey, 2003-2008
' REV: 1.0
Set WS = WScript.CreateObject("WScript.Shell")
Set WN = WScript.CreateObject("WScript.Network")
Set fso = CreateObject("Scripting.FileSystemObject")
Const bfnRoot = "D:\data\"   'NEEDS SUFFIX '\'
'******************************************************************************
' set gnu pg settings folder
'WS.RegWrite "HKEY_CURRENT_USER\Software\GNU\GnuPG\HomeDir","D:\data\gnupg","REG_SZ"
' kill the search dog
WS.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState\Use Search Asst","no","REG_SZ"
'disable skype supernode
WS.RegWrite "HKLM\SOFTWARE\Policies\Skype\Phone\DisableSupernode",1,"REG_DWORD"
'******************************************************************************
WS.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\ForceAutoLogon","1","REG_SZ"
'******************************************************************************
'turn on num-lock for logon window
WS.RegWrite "HKEY_USERS\.Default\Control Panel\Keyboard\InitialKeyboardIndicators","2","REG_SZ" 


'******************************************************************************
' set OE store folder
WS.RegWrite "HKCU\Identities" & WS.RegRead("HKCU\Identities\Default User ID") & "\Software\Microsoft\Outlook Express\5.0",bfnRoot&"\OE Mail Store","REG_SZ"

on error resume next
' Remove Shared Docs from My Computer
WS.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\DelegateFolders\{59031a47-3f72-44a7-89c5-5595fe6b30ee}\"
on error goto 0
'******************************************************************************
commondesktop=bfnroot & WN.ComputerName & "\AllUsers\Desktop"
commonstartup=bfnroot & WN.ComputerName & "\AllUsers\Start Menu\Programs\Startup"
commonprograms=bfnroot & WN.ComputerName & "\AllUsers\Start Menu\Programs"
commonstartmenu=bfnroot & WN.ComputerName & "\AllUsers\Start Menu"
'
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Desktop",commondesktop,"REG_EXPAND_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Start Menu",commonstartmenu,"REG_EXPAND_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Startup",commonstartup,"REG_EXPAND_SZ"
WS.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Programs",commonprograms,"REG_EXPAND_SZ"
'end if        
'******************************************************************************
'set app paths - (for start>run)
'const app = "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\"
'
'WS.RegWrite app & "writer.exe\","C:\Program Files\OpenOffice.org 2.3\program\swriter.exe","REG_SZ"
'WS.RegWrite app & "thunder.exe\","C:\Program Files\Mozilla Thunderbird\thunderbird.exe","REG_SZ"
'WS.RegWrite app & "excel.exe\","C:\Program Files\OpenOffice.org 2.3\program\scalc.exe","REG_SZ"
' USE SHORTCUTS IN C:\windows directory instead.

'******************************************************************************
'set shell folder paths
admintools = bfnroot & "%COMPUTERNAME%\AllUsers\Start Menu\Programs\Administrative Tools"
Personal = "F:\documents"
Favorites=bfnRoot & "FAVORITES"
My_Music=personal & "\music"
My_Pictures=personal & "\pictures"
My_Video=personal & "\video"
'
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Administrative Tools",admintools,"REG_EXPAND_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal",Personal,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites",FAVORITES,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Music",My_Music,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures",My_Pictures,"REG_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Video",My_Video,"REG_SZ"
'
startup=bfnroot &  WN.ComputerName & "\%USERNAME%\Start Menu\Programs\Startup"
programs=bfnroot &  WN.ComputerName & "\%USERNAME%\Start Menu\Programs"
startmenu=bfnroot &  WN.ComputerName & "\%USERNAME%\Start Menu"
desktop = bfnRoot  &  WN.ComputerName & "\%USERNAME%\desktop"
'
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Desktop",Desktop,"REG_EXPAND_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\startup",startup,"REG_EXPAND_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\programs",programs,"REG_EXPAND_SZ"
WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Start menu",startmenu,"REG_EXPAND_SZ"
'******************************************************************************
' delete unneeded files
on error resume next
QLaunch="C:\Documents and Settings\"&WN.UserName&"\Application Data\Microsoft\Internet Explorer\Quick Launch\"
'FSO.DeleteFile(QLaunch & "Launch Outlook Express.lnk")
on error goto 0
