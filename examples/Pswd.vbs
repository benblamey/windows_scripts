' Pswd.vbs
' Demonstrates the InputBox object
'-----------------------------------------------------------------

Dim dlg, strText

Set dlg = CreateObject("VbsInput.InputBox")

dlg.Title = "My Application"
dlg.Label = "Type in your password"
dlg.Icon = "shell32.dll,44"
dlg.Password = true

strText = dlg.Show
If strText <> "" Then
  MsgBox strText
End If




