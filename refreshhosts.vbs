option explicit
dim loc
loc=SetLocale("en-gb")
dim WS, WN, strPassB, fso,f,f2,strLine
Set WS = WScript.CreateObject("WScript.Shell")
Set WN = WScript.CreateObject("WScript.Network")
Set fso = CreateObject("Scripting.FileSystemObject")

'strPassB = WS.RegRead("HKLM\Software\BFN\PassAdmin")
Const bfnRoot = "D:\data\"
Const hostsfile="C:\windows\system32\drivers\etc\hosts"

WS.Run "C:\progra~1\wget\wget.exe http://www.mvps.org/winhelp2002/hosts.txt -O C:\windows\system32\drivers\etc\hosts", ,-1

Set f = fso.OpenTextFile(hostsfile, 8, -1)
Set f2 = fso.OpenTextFile(bfnRoot & "hosts\append.txt", 1, -1)
do while 1=1
	strLine = f2.ReadLine 
	f.WriteLine strLine
	if strLine = "#EOF" then exit do
loop 
f.Close
f2.Close
