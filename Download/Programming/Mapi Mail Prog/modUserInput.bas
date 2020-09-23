Attribute VB_Name = "modUserInput"
Public vDate As Date
Public vTXT As String
Public vCancel As Boolean
Public vMsgID As Long
Public vMsgDate  As Date

Function VerifyListDate(pForm As Form, txt As String, msgID As Long, msgDate As Date) As Date
vTXT = txt
vMsgID = msgID
vMsgDate = msgDate
vCancel = True
frmVerifyDate.Show 1, pForm
If vCancel = False Then
VerifyListDate = vDate
Else
ret& = MsgBox("No date specified.Continue using Today's date ?", vbYesNo + vbDefaultButton1 + vbQuestion, "Verify date")
    If ret& = vbYes Then
    vDate = Date
    Else
    vDate = VerifyListDate
    VerifyListDate = vDate
    End If
End If






End Function


