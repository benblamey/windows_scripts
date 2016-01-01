Set WS = WScript.CreateObject("WScript.Shell")
Ws.RegWrite "HKLM\Software\BFN\SMART_ERROR","123456","REG_SZ"
WS.Run "shutdown.exe -s -t 300 -c "&chr(34)&"WARNING! Critical hard drive malfunction detected. automatic shutdown initiated. Contact Ben for assistance. Hard drive failuare predicted. Catastrophic data loss likely. Save all work immediately. ***********"&chr(34), , 0
