Set WS = WScript.CreateObject("WScript.Shell")
strPassB = WS.RegRead("HKLM\Software\BFN\PassAdmin")
WS.Run "psexec.exe -u admin -p "&strPassB&" -e -d "&chr(34)&"C:\Program Files\Lexmark X1100 Series\lxbkbmgr.exe"&Chr(34), , 0
