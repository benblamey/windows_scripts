Set WS = WScript.CreateObject("WScript.Shell")
strPassB = WS.RegRead("HKLM\Software\BFN\PassAdmin")
WS.Run "psexec.exe -u admin -p tuc4568@q -e -d "&chr(34)&"C:\Program Files\Lexmark X1100 Series\lxbkaiox.exe"&chr(34)
