Set WS = WScript.CreateObject("WScript.Shell")
strPassB = WS.RegRead("HKLM\Software\BFN\strPassB")
WS.Run "psexec.exe -u admin -p tuc4568@q -d "&chr(34)&"C:\Program Files\Lexmark X1100 Series\lxbkvb.exe"