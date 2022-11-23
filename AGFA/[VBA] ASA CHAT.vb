

Sub ASApctCHAT(Optional HideMe As String)
'
' ASApctCHAT Macro
'

'
Line1:
ActiveSheet.Range("L2") = InputBox("Enter number of seconds", "SET ASA THRESHOLD")
    If Range("L2") = "" Then
        MsgBox "Please enter data", vbCritical + vbOKOnly, "Error"
        GoTo Line1
        Else
    MsgBox "ASA 85% within " & Range("L2") & """" + VbMsgBoxCenter
    End If