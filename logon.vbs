dim loc
loc=SetLocale("en-gb")

'************************************
' (c) Ben Blamey, 2003-2008
' REV: 1.3
Set WS = WScript.CreateObject("WScript.Shell")
Set WN = WScript.CreateObject("WScript.Network")
Set fso = CreateObject("Scripting.FileSystemObject")
'****************
'const avgdelay= 1 'number of days between updating avg
const PageDefrag= 5 'number of days between defrag pagefile
const DefragDelay= 5 'number of days between defrag pagefile
const HostsDelay= 10 'number of days between refreshing hostsfile
'Const viruscheckinterval= 7   'days between doing a full virus scan
Const AUTOLOGONSTATUS="1"   '1=autologon enabled, 0=disabled
const WinLogRoot = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\"
CONST bfnRoot = "D:\data\"
const pagedefragexecpath="C:\Progra~1\pasteonlyapps\sysinternals\pagedfrg.exe -o -t 1"
const tucoysecalbackup= 1 'number of days tucoyse cal update
'***************

UN=UCASE(WN.UserName)
CN=UCASE(WN.ComputerName)
'***************
'avguppath = "C:\Progra~1\Grisoft\AVG7\avginet.exe /SCHED="
'avgcheckpath ="C:\Progra~1\Grisoft\AVG7\avgw.exe /TEST=2"
'***************
on error resume next
	fso.DeleteFile "C:\Progra~1\Opera\profile\sessions\*.win", -1
	fso.DeleteFile "C:\Progra~1\Opera\profile\sessions\*.win.bak", -1
	fso.DeleteFile bfnRoot&"opera\sessions\*.win", -1
	fso.DeleteFile bfnRoot&"opera\sessions\*.win.bak", -1
	fso.CopyFile bfnRoot&"opera\sessions\autosave.backup", "D:\data\opera\sessions\autosave.win"
on error goto 0
'***************
strPassUser = WS.RegRead("HKLM\Software\BFN\PassUser")

if UN="USER" then
	WS.RegWrite WinLogRoot&"DefaultUserName","USER","REG_SZ"
	WS.RegWrite WinLogRoot&"AutoAdminLogon",AUTOLOGONSTATUS,"REG_SZ"
	WS.RegWrite WinLogRoot&"DefaultPassword",strPassUser,"REG_SZ"
	Ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden", 2, "REG_DWORD" 'hide hidden files for user
end if
if UN="ADMINISTRATOR" then
	WS.RegWrite WinLogRoot&"DefaultUserName","ADMINISTRATOR","REG_SZ"
	WS.RegWrite WinLogRoot&"AutoAdminLogon",0,"REG_SZ"
	WS.RegWrite WinLogRoot&"DefaultPassword","","REG_SZ"
end if
'************************************************************
if CN="TUCOYSE" then
	' Tucoyse Cal
	dater = RegRead("HKLM\Software\BFN\tuccal")
	dater = DateDiff("d", Date, dater)
	if dater =< -tucoysecalbackup then 
		WS.Run("python D:\DOCS\Tucoyse\apps\tucoysecal\tucoyse.py silent")
		'WS.Run "wscript D:\DOCS\Tucoyse\apps\tucoysecal\cal.vbs silent", , -1
		WS.RegWrite "HKLM\Software\BFN\tuccal", Date, "REG_SZ"
	end if 
 
' Time in mS before an application is considered to be not responding, on logoff/shutdown
WS.RegWrite "HKCU\Control Panel\Desktop\WaitToKillAppTimeout","20000","REG_SZ"

' After this time, the process is ended automatically=1. Confirmation to end=2.
WS.RegWrite "HKCU\Control Panel\Desktop\AutoEndTasks","1","REG_SZ"

end if
'************************************************************
'AVG Update
'	dater = RegRead("HKLM\Software\BFN\avgup")
'	dater = DateDiff("d", Date, dater)
'	if dater =< -AVGDelay then 
'		WS.Run avguppath, , -1
'		WS.RegWrite "HKLM\Software\BFN\avgup", Date, "REG_SZ"
'	end if
'************************************************************
'file defrag
	dater = RegRead("HKLM\Software\BFN\Defrag")
	dater = DateDiff("d", Date, dater)
	if dater =< -DefragDelay  then
		WS.Run "defrag C: -v", , -1
		WS.Run "defrag D: -v", , -1
		WS.RegWrite "HKLM\Software\BFN\Defrag", Date, "REG_SZ"
	end if
'************************************************************	
'HOSTS
	if CN="BOMBADIL" or CN="BEN" or CN="TUCOYSE" then
		dater = RegRead("HKLM\Software\BFN\Hosts")
		dater = DateDiff("d", Date, dater)
		if dater =< -HostsDelay  then
			WS.Run "wscript "&bfnroot&"scripts\refreshhosts.vbs"
			WS.RegWrite "HKLM\Software\BFN\Hosts", Date, "REG_SZ"
		end if
	end if
'************************************************************	
' SMART error flag
	dater = RegRead("HKLM\Software\BFN\SMART_ERROR")
	if dater="123456" then
		WScript.Echo("WARNING! Critical hard drive S.M.A.R.T. malfunction detected. automatic shutdown initiated. Contact Ben for assistance. Hard drive failuare predicted. Catastrophic data loss likely. Save all work immediately & shutdown. Registry flag requires manual deletion.")
'		WS.RegWrite "HKLM\Software\BFN\SMART_ERROR","0","REG_SZ"
	end if
'************************************************************	
' Virus Check
'	dater = RegRead("HKLM\Software\BFN\VirusCheck")
'	dater = DateDiff("d", Date, dater)
'	if dater =< -viruscheckinterval then 
'		WS.Run avgcheckpath
'		WS.RegWrite "HKLM\Software\BFN\VirusCheck", Date, "REG_SZ"
'	end if  
'************************************************************	
' Page Defrag
	dater = RegRead("HKLM\Software\BFN\PageDefrag")
	dater = DateDiff("d", Date, dater)
	if dater =< -PageDefrag then 
		WS.Run pagedefragexecpath, ,-1
		WS.RegWrite "HKLM\Software\BFN\PageDefrag", Date, "REG_SZ"
	end if 
'************************************************************	
'if CN="BOMBADIL" then 
'	passBoatClub = WS.RegRead("HKLM\Software\BFN\boatclub")
'	passNetware = WS.RegRead("HKLM\Software\BFN\Netware")
'
'	on error resume next	
'		WN.RemoveNetworkDrive "Y:"	
'		WN.RemoveNetworkDrive "W:"	
'	on error GoTo 0
'	Wscript.Sleep(6000)
'
'	WN.MapNetworkDrive "Y:", "\\chufilestore\A_D\brhb2", 0, "brhb2", passNetware
'	Wscript.Sleep(2000)
'	WN.MapNetworkDrive "W:", "\\131.111.131.24\A_D\chuboat", 0, "chuboat", passBoatClub
'end if
'************************************************************	
'if CN="BEN" or CN="BOMBADIL" then
'	WS.Run("attrib +H D:\Docs\*.???~ /S")  'Marks gvim backups as hidden. (and lyx) (linux programs in general use this)
'	WS.Run("attrib +H D:\Data\*.???~ /S")  'Marks gvim backups as hidden.
'end if

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
Function RegRead(ByVal sRegValue) 
On Error Resume Next
RegRead = WS.RegRead(sRegValue)
On Error Goto 0
' If the value does not exist, error is raised
If Err Then
	RegRead = "01/01/1986"
	Err.clear
End If
If VarType(RegRead) < vbArray then
	If RegRead = sRegValue Then RegRead = ""
End If
End Function
