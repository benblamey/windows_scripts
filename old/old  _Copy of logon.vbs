'option explicit
dim WS,WN,fso,UN,dater,f,f2,strLine,strPassUser,strPassB,response,reeturn,daterr,returnn,silent,objArgs,strPassAdmin
Set WS = WScript.CreateObject("WScript.Shell")
Set WN = WScript.CreateObject("WScript.Network")
Set fso = CreateObject("Scripting.FileSystemObject")
'******
const avgdelay= 1 'number of days between updating avg
const PageDefrag= 9 'number of days between defrag pagefile
const DefragDelay= 9 'number of days between defrag
const BackupDelay= 9 'number of days between backups
const backupselection="D:\backup\selection.bks"
const incselection="D:\backup\selection.bks"
const backuptgtfolder = "D:\backup\scheduled\"
const avguppath = "C:\Progra~1\Grisoft\AVG7\avginet.exe /SCHED="
Const AUTOLOGONSTATUS="1"
const idleearly = 2   ' how many days early to do when running idle
const WinLogRoot = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\"
strPassUser = WS.RegRead("HKLM\Software\BFN\PassUser")
strPassB = WS.RegRead("HKLM\Software\BFN\PassAdmin")
strPassAdmin = WS.RegRead("HKLM\Software\BFN\PassAdmin")
UN=UCASE(WN.UserName)
CN =UCASE(WN.ComputerName="TUCOYSE")
'***************
if CN="BOMBADIL" then bfnRoot = "D:\data\"
if CN="BEN" then bfnRoot = "D:\data\"
if CN="TUCOYSE" then bfnRoot = "D:\bfn_data\"
'***************
silent=0
Set objArgs = WScript.Arguments
if objArgs.Count = "1" then silent=1
if silent = 0 then
	'***************
	on error resume next
	fso.DeleteFile "C:\Progra~1\Opera\profile\sessions\autosave.win", -1
	fso.DeleteFile "C:\Progra~1\Opera\profile\sessions\autosave.win.bak", -1
	on error goto 0
	'***************
	if UN="BEN" or UN="ADMIN" then
		WS.RegWrite WinLogRoot&"DefaultUserName","admin","REG_SZ"
		WS.RegWrite WinLogRoot&"AutoAdminLogon",AUTOLOGONSTATUS,"REG_SZ"
		WS.RegWrite WinLogRoot&"DefaultPassword",strPassB,"REG_SZ"
		'show hidden files for admin
		WS.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden", 1, "REG_DWORD"
	end if
	if UN="USER" then
		WS.RegWrite WinLogRoot&"DefaultUserName","USER","REG_SZ"
		WS.RegWrite WinLogRoot&"AutoAdminLogon",AUTOLOGONSTATUS,"REG_SZ"
		WS.RegWrite WinLogRoot&"DefaultPassword",strPassUser,"REG_SZ"
		Ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden", 2, "REG_DWORD" 'hide hidden files for user
	end if
	if UN="ADMINISTRATOR" then
		WS.RegWrite WinLogRoot&"DefaultUserName","ADMINISTRATOR","REG_SZ"
		WS.RegWrite WinLogRoot&"AutoAdminLogon",AUTOLOGONSTATUS,"REG_SZ"
		WS.RegWrite WinLogRoot&"DefaultPassword","39f£ch359^","REG_SZ"
	end if
	'***************
	if CN="TUCOYSE" then
		'***************
		' Tucoyse Cal
		const tucoysecalbackup= 1 'number of days tucoyse cal update
		dater = RegRead("HKLM\Software\BFN\tuccal")
		dater = DateDiff("d", Date, dater)
		if dater =< -tucoysecalbackup then 
			WS.Run "wscript D:\DOCS\Tucoyse\apps\tucoysecal\cal.vbs silent", , -1
			WS.RegWrite "HKLM\Software\BFN\tuccal", Date, "REG_SZ"
		end if 
	else
		'***************
		'AVG Update
		dater = RegRead("HKLM\Software\BFN\avgup")
		dater = DateDiff("d", Date, dater)
		if dater =< -AVGDelay then 
			WS.Run avguppath, , -1
			WS.RegWrite "HKLM\Software\BFN\avgup", Date, "REG_SZ"
		end if
	end if
end if
'***************
'file defrag
dater = RegRead("HKLM\Software\BFN\Defrag")
dater = DateDiff("d", Date, dater)
if (dater =< -DefragDelay and silent=0) or (silent=1 and dater =< (-DefragDelay)+IdleEarly) then 
		WS.Run "psexec -u admin -p " & strPassB & " defrag C: -v", , -1
		WS.Run "psexec -u admin -p " & strPassB & " defrag D: -v", , -1
		WS.RegWrite "HKLM\Software\BFN\Defrag", Date, "REG_SZ"
		bfnend()
end if
'***************
' backup
dater = RegRead("HKLM\Software\BFN\Backup")
dater = DateDiff("d", Date, dater)
if (dater =< -BackupDelay and silent=0) or (silent=1 and dater =< (-BackupDelay)+IdleEarly) then 
		On Error Resume Next
			FSO.DeleteFile (backuptgtfolder&"\_old_full.bkf")
			FSO.MoveFile backuptgtfolder&"\_new_full.bkf", backuptgtfolder&"\_old_full.bkf" 
		On Error GoTo 0
		Wscript.Echo("1")
		reeturn = WS.Run ("psexec -u admin -p " & strPassB & " ntbackup.exe backup " & chr(34) & "@" & incselection & chr(34) & " /J " & chr(34) & "weekly backup" & int(Timer) & chr(34) & " /F " & chr(34) & backuptgtfolder & DatePart("YYYY", date) & "-" & DatePart("m",date) & "-" & DatePart("d",date) &"-"& int(Timer) & "-inc.bkf" & chr(34) & " /D " & chr(34) & "Set created " & Date & chr(34) & " /V:no /R:no /L:n /M incremental /RS:no /HC:off /SNAP:off", , -1)
		if reeturn<>0 then Wscript.Echo("Incremental Backup Terminated with an error code of: " & reeturn & ". Please report this to Ben.")
		Wscript.Echo("2")
		reeturn = WS.Run ("psexec -u admin -p " & strPassB & " ntbackup.exe backup " & chr(34) & "@" & backupselection & chr(34) & " /J " & chr(34) & "weekly backup" & int(Timer) & chr(34) & " /F " & chr(34) & backuptgtfolder & "_new_full.bkf" & chr(34) & " /D " & chr(34) & "Set created " & Date & chr(34) & " /V:no /R:no /L:n /M normal /RS:no /HC:off /SNAP:off", , -1)
		if reeturn<>0 then Wscript.Echo("Normal Backup Terminated with an error code of: " & reeturn & ". Please report this to Ben.")
		WS.RegWrite "HKLM\Software\BFN\Backup", Date, "REG_SZ"
		bfnend()
		Wscript.Echo("3")
end if
'***************
' Page Defrag
dater = RegRead("HKLM\Software\BFN\PageDefrag")
dater = DateDiff("d", Date, dater)
if (dater =< -PageDefrag and silent=0) or (silent=1 and dater =< (-PageDefrag)+IdleEarly) then 
	WS.Run "psexec -u admin -p " & strPassB & " pagedfrg.exe -o -t 1", ,-1
	WS.RegWrite "HKLM\Software\BFN\PageDefrag", Date, "REG_SZ"
end if  
'##########################################
if silent=0 then 
	WS.Run "psexec -u admin -p " & strPassB & " -d C:\Progra~1\DellSupport\DSAgnt.exe /startup", ,-1
end if
if Hour(time) = 2 and silent = 1 then WS.Run("shutdown -s -t 600")
'##########################################
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
'##########################################
Function bfnend()
	if Hour(time) = 2 and silent = 1 then 
'		if CN="BOMBADIL" or CN="BEN" then
'			WS.Run("sendemail -f robert@blamey1585.freeserve.co.uk -t brhb2@cam.co.uk -u Last night -m You left me on last night and I turned myself off. your Computer. xx -s smtp.orangehome.co.uk:25")
'		end if	
'		if CN="TUCOYSE" then
'			WS.Run("sendemail -f robert@blamey1585.freeserve.co.uk -t penny@tucoyse.co.uk -u Last night -m You left me on last night and I turned myself off. your Computer. xx -s smtp.orangehome.co.uk:25")
'		end if	
		WS.Run("shutdown -s -t 600")
	end if
	if silent=1 then wscript.quit
End Function
