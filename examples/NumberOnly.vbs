Dim dlg, strText

Set dlg = CreateObject("VbsInput.InputBox")

dlg.Title = "My Application"
dlg.Label = "Enter your age"
dlg.Icon = "shell32.dll,54"
dlg.NumberOnly = True
dlg.Align="right"

strText = dlg.Show()
If strText <> "" Then
  MsgBox "Your age is " & strText
End If




