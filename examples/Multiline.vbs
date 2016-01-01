Dim dlg, strText

Set dlg = CreateObject("VbsInput.InputBox")

dlg.Title = "My Application"
dlg.Label = "Enter your comments (max 250 characters)"
dlg.Icon = "shell32.dll,54"
dlg.multiline = True

' You could even set this!
'dlg.Align="right"

strText = dlg.Show()
If strText <> "" Then
  MsgBox strText
End If




