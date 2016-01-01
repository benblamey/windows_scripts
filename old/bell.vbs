on error resume next
' 0 subject
' 1 sender      name name <email@domain>
' 2 recp list
' 3 12345??
' 4 folder
' 5 type (IMAP)
' 6 account
' 7 body

Set objArgs = WScript.Arguments
Set objVoice = CreateObject("SAPI.SpVoice") 
Set objVoice.Voice = objVoice.GetVoices("Name=Microsoft Sam").Item(0) 
Set WS = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'*************' dont beep for spam/junk
if UCASE(objArgs(4))="SPAM" then WScript.Quit
if UCASE(objArgs(4))="JUNK" then WScript.Quit


' -----------------SUBJECT-----------------

subject = objArgs(0)


'*************' strip off RE:, FWD: etc.
'if UCASE(mid(subject,1,4)) = "RE: " then subject = "reply" &mid(subject,4)
'if UCASE(mid(subject,1,4)) = "FW: " then subject = "forward" &mid(subject,4)
'if UCASE(mid(subject,1,5)) = "FWD: " then subject = "forward" & mid(subject,5)

'if UCASE(mid(subject,1,3)) = "RE:" then subject = "reply " &mid(subject,4)
'if UCASE(mid(subject,1,3)) = "FW:" then subject = "forward " &mid(subject,4)
'if UCASE(mid(subject,1,4)) = "FWD:" then subject = "forward " & mid(subject,5)

'*************'cut out punctuation from subject


For counter = 1 To len(subject)
	if mid(subject,counter,1)<>"*" and mid(subject,counter,1)<>":" and mid(subject,counter,1)<>"!" and mid(subject,counter,1)<>"." and mid(subject,counter,1)<>"-" and mid(name,counter,1)<>"_" and mid(name,counter,1)<>"@" then 
		subject2 = subject2 & mid(subject,counter,1)
	else
		subject2= subject2 & " "
	end if
next
subject=subject2



'-------------------NAME-------------------
name2 = objArgs(1)
if mid(objArgs(1),1,1)="\" then name2 = mid(objArgs(1),2) & " " & objArgs(2)

'*************' chop name at  '<'
For counter = 1 To len(name2)
	if asc(mid(name2,counter,1))=60 or mid(name2,counter,1)="\" then Exit For    '     ('<' character)
Next
name=Mid(name2, 1, Counter-1)


'************** remove .'s from name
name2=""
For counter = 1 To len(name)
	if mid(name,counter,1)<>"." and mid(name,counter,1)<>"_" and mid(name,counter,1)<>"@" then 
		name2 = name2 & mid(name,counter,1)
	else
		name2=name2&" "
	end if
Next
name=name2

'---------------SPEAK------------

objVoice.Speak "Hi Ben, you have a new e-mail from " & name & ".  the Subject is   " & subject
Wscript.Sleep(1000)
Ws.Run "cmd /c echo ", 7

'************* check output for decoding purposes
for counter = 0 to 4
line = line &"*" &counter&"*" & objArgs(counter)
next
line = line &"-/-"&name&"-/-"&subject


'Set f = fso.OpenTextFile("D:\data\bombadil\user\desktop\bell.txt", 8, True)   '(for appending)
'f.WriteLine line
'f.close
