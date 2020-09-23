Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Type CodeEntry
    ID As Long 'eMail message ID
    Title As String 'Title of code entry
    Text As String 'Description
    http As String 'Location
    Category As String
    Level As String
    Compatibility As String
    DateSubmitted As Date
    Accessed As Long
    ListDate As Date
    CodeOfTheDay As Boolean
End Type
Public CE As CodeEntry

Public Type RTFIndexStyle
    Start As Long
    End As Long
    Style As RTFStyle
End Type
Public rtIndex() As RTFIndexStyle
Public rtIndexCount As Long


Public Enum RTFStyle
    stlTitle = 1
    stlField = 2
    stltext = 3
    stlhttp = 4
    stlComment = 5
End Enum


Public frmMessageTitle As String
Public frmMessageMessage As String
Public frmMessageCheckVisible As Boolean
Public frmMessageCheckKey As String
Public frmMessageButtonCaption As String
Public frmMessageCheckCaption As String

Public Const LB_FINDSTRING = &H18F
Public Const LB_ERR = (-1)
Public Function FindStringExact(listHwnd As Long, ByVal IndexStart As Long, ByVal srcString As String) As Long
ret& = SendMessage(listHwnd, LB_FINDSTRING, IndexStart, ByVal srcString)
If ret& = LB_ERR Then
FindStringExact = -1
Else
FindStringExact = ret&
End If

End Function


Public Function checkIfEmail(eMail As String) As Boolean
    Dim i As Integer
    Dim char As String
    Dim c() As String
    'checks if the string has the standard e
    '     mail pattern:


    If Not eMail Like "*@*.*" Then
        checkIfEmail = False
        Exit Function
    End If
    'splits the email-string with a "." deli
    '     meter and returns the subtring in the c-
    '     string array
    c = Split(eMail, ".", -1, vbBinaryCompare)
    'checks if the last substring has a leng
    '     th of either 2 or 3


    If Not Len(c(UBound(c))) = 3 And Not Len(c(UBound(c))) = 2 Then
        checkIfEmail = False
        Exit Function
    End If
    'steps through the last substring to see
    '     if it contains anything else unless char
    '     acters from a to z


    For i = 1 To Len(c(UBound(c))) Step 1
        char = Mid(c(UBound(c)), i, 1)


        If Not (LCase(char) <= Chr(122)) Or Not (LCase(char) >= Chr(97)) Then
            checkIfEmail = False
            Exit Function
        End If
    Next i
    'steps through the whole email string to
    '     see if it contains any special character
    '     s:


    For i = 1 To Len(eMail) Step 1
        char = Mid(eMail, i, 1)
        If (LCase(char) <= Chr(122) And LCase(char) >= Chr(97)) _
        Or (char >= Chr(48) And char <= Chr(57)) _
        Or (char = ".") _
        Or (char = "@") _
        Or (char = "-") _
        Or (char = "_") Then
        checkIfEmail = True
    Else
        checkIfEmail = False
        Exit Function
    End If
Next i
End Function

Function TestSQL(SQL As String) As Boolean
'Checks if a SQL string is valid

On Error GoTo errSQL1
Dim testDB As Database
Dim testSet As Recordset
Set testDB = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set testSet = testDB.OpenRecordset(SQL, dbOpenDynaset)
testSet.Close
testDB.Close
Set testDB = Nothing
Set testSet = Nothing
TestSQL = True
Exit Function


errSQL1:
TestSQL = False
MsgBox Error
Exit Function
Resume Next


End Function

Sub DisplayImage(img As Image, resID As Long)
img.Picture = LoadResPicture(resID, vbResBitmap)
img.Refresh
End Sub

Sub ClearEmptyLines(txt As TextBox)
Dim msgLineCount As Long
Dim lineMSG As String
Dim tempSTR As String

msgLineCount = SendMessage(txt.hwnd, EM_GETLINECOUNT, 0, 0)
For i% = 0 To msgLineCount - 1
lineMSG = GetLine(txt.hwnd, i%)
    If Len(lineMSG) > 0 Then
        If Left$(lineMSG, 2) <> vbCrLf Then
        If i% > 0 Then
        tempSTR = tempSTR & vbCrLf & lineMSG
        Else
        tempSTR = lineMSG
        End If
        End If
    End If
Next i%
txt.Text = tempSTR


End Sub

Function GetCodeEntriesCount(txtHwnd As Long, note As String, LineNo As Integer) As Integer
Dim cLine As String
Dim count As Integer
Dim lineNumber As Integer

For i% = 0 To GettxtLineCount(txtHwnd) - 1
    cLine = GetLine(txtHwnd, i%)
    LineNo = i%
    If Left$(cLine, 5) = "*****" Or Left$(cLine, 5) = "=====" Then Exit For
    For d% = 1 To 4
        If Mid$(cLine, d%, 1) = ")" Then
        If IsNumeric(Left(cLine, d% - 1)) Then lineNumber = CInt(Left(cLine, d% - 1))
        count = lineNumber
        Exit For
        End If
    Next d%
Next i%
GetCodeEntriesCount = count
    
End Function

Public Function GetLine(ByVal emHWND As Long, ByVal lineNumber As Integer) As String
If lineNumber > SendMessage(emHWND, EM_GETLINECOUNT, 0, 0) - 1 Then
GetLine = ""
Exit Function
End If

Const MAX_CHAR_PER_LINE = 80
' This function fills the buffer with a line of text
      ' specified by LineNumber from the text-box control.
      ' The first line starts at zero.
      byteLo% = MAX_CHAR_PER_LINE And (255)  '[changed 5/15/92]
      byteHi% = Int(MAX_CHAR_PER_LINE / 256) '[changed 5/15/92]
      Buffer$ = Chr$(byteLo%) + Chr$(byteHi%) + Space$(MAX_CHAR_PER_LINE - 2)
      ' [Above line changed 5/15/92 to correct problem.]
       X = SendMessageAsString(emHWND, EM_GETLINE, lineNumber, Buffer$)
      GetLine = Left$(Buffer$, X)
      

    
End Function

Function GettxtLineCount(txtHwnd As Long) As Long
GettxtLineCount = SendMessage(txtHwnd, EM_GETLINECOUNT, 0, 0)

End Function


Function RemoveLinesUntilString(txt As TextBox, x1 As String, x2 As String) As Long
Dim lb As Integer
lb% = Len(vbCrLf)
Do While GettxtLineCount(txt.hwnd) > 0
cLine$ = UCase(GetLine(txt.hwnd, 0))

If InStr(1, cLine$, UCase(x1)) = 0 And InStr(1, cLine$, UCase(x2)) = 0 Then
txt.Text = Right$(txt.Text, (Len(txt) - (Len(cLine$) + lb%)))
Else
txt.Text = Right$(txt.Text, (Len(txt) - (Len(cLine$) + Len(lb))))
RemoveLinesUntilString = GettxtLineCount(txt.hwnd)
Exit Do
End If
Loop

End Function


