Set WS = WScript.CreateObject("WScript.Shell")
strPassB = WS.RegRead("HKLM\Software\BFN\PassAdmin")
WS.Run "psexec.exe -u admin -p "&strPassB&" -e -d "&chr(34)&"C:\WINDOWS\system32\spool\drivers\w32x86\3\LXBKPSWX.EXE"&chr(34)&" /M=Lexmark X1100 Series /T=100"
