Set WS = WScript.CreateObject("WScript.Shell")
strPassB = WS.RegRead("HKLM\Software\BFN\strPassB")
WS.Run "psexec.exe -u admin -p "&strPassB&" -e -d "&chr(34)&"C:\Program Files\Mozilla Thunderbird\thunderbird.exe"&Chr(34), , 0
