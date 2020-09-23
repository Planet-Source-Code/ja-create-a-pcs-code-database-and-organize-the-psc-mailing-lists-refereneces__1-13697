VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind_eMails 
   Caption         =   "Find eMail addresses"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   Icon            =   "frmFind_eMails.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "&Remove"
      Height          =   345
      Left            =   4485
      TabIndex        =   12
      Top             =   2070
      Width           =   945
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Remove &All"
      Height          =   345
      Left            =   4485
      TabIndex        =   11
      Top             =   2490
      Width           =   945
   End
   Begin VB.CommandButton Command7 
      Caption         =   "View msg"
      Height          =   345
      Left            =   4485
      TabIndex        =   10
      Top             =   2895
      Width           =   945
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Close"
      Height          =   420
      Left            =   4515
      TabIndex        =   9
      Top             =   1260
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Save"
      Height          =   420
      Left            =   4500
      TabIndex        =   8
      Top             =   735
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove &All"
      Height          =   345
      Left            =   3225
      TabIndex        =   7
      Top             =   1135
      Width           =   945
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Remove"
      Height          =   345
      Left            =   3225
      TabIndex        =   6
      Top             =   710
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Height          =   345
      Left            =   3225
      TabIndex        =   5
      Top             =   285
      Width           =   945
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   60
      TabIndex        =   3
      Top             =   285
      Width           =   3105
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&tart"
      Height          =   420
      Left            =   4500
      TabIndex        =   2
      Top             =   225
      Width           =   1500
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   210
      Left            =   1515
      TabIndex        =   1
      Top             =   1755
      Visible         =   0   'False
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.ListBox List1 
      Height          =   1635
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   2025
      Width           =   4365
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   2475
      Top             =   3165
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3165
      Top             =   3165
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Hit ""Start"" to begin"
      Height          =   210
      Left            =   75
      TabIndex        =   13
      Top             =   1755
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Keywords or Keyfhrases to look for"
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   60
      Width           =   2865
   End
End
Attribute VB_Name = "frmFind_eMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public msgNote_tmp As String


Sub AddToList(Add As String, index As Long, note As String)
If FindStringExact(List1.hwnd, 0, Add & vbNullString) = -1 Then
List1.AddItem Add
List1.Selected(List1.NewIndex) = True
List1.ItemData(List1.NewIndex) = index
ReDim Preserve cValid_eMails(List1.ListIndex)
cValid_eMails(List1.ListIndex).eMail = Add
cValid_eMails(List1.ListIndex).index = index
cValid_eMails(List1.ListIndex).msgSource = note
List2.Refresh
End If

End Sub

Function Valid_eMail(note As String) As Boolean
For i% = 0 To List2.ListCount - 1
t$ = List2.List(i%)
    If InStr(1, note, t$) > 0 Then
    Valid_eMail = True
    Exit Function
    End If
Next i%
Valid_eMail = False

    
End Function

Private Sub Command1_Click()
List1.Clear
ReDim cValid_eMails(0)
If List2.ListCount < 0 Then
MsgBox "You must include at least one keyword/phrase in the list to look for"
Exit Sub
End If

List1.Enabled = False
MousePointer = 11
MAPISession1.SignOn
Label2.Caption = "Fetching messages"
Refresh
MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.Fetch
If MAPIMessages1.msgCount = 0 Then: GoTo Exiting
pb1.Min = 0
pb1.Max = MAPIMessages1.msgCount - 1
pb1.Value = pb1.Min
pb1.Visible = True
List1.Clear
Refresh
Label2.Caption = "Searching..."
Refresh
For i% = 0 To MAPIMessages1.msgCount - 1
pb1.Value = i%
MAPIMessages1.MsgIndex = i%
    If Valid_eMail(MAPIMessages1.MsgNoteText) = True Then
    AddToList MAPIMessages1.MsgOrigAddress, MAPIMessages1.MsgIndex, MAPIMessages1.MsgNoteText
    End If
Next i%


Exiting:
MousePointer = 0
MAPISession1.SignOff
Label2.Caption = "Finished..."
pb1.Visible = False
List1.Enabled = True
End Sub

Private Sub Command2_Click()
d$ = InputBox("Enter a word or phrase to include in the 'Search for list'")
If d$ <> "" Then List2.AddItem d$


End Sub


Private Sub Command3_Click()
If List2.ListIndex > -1 And List2.ListIndex < List2.ListCount Then
List2.RemoveItem List2.ListIndex
End If

End Sub


Private Sub Command4_Click()
List2.Clear

End Sub


Private Sub Command5_Click()
For i% = 0 To List1.ListCount - 1
X$ = List1.List(i%)
    If List1.Selected(i%) = True Then
        If FindStringExact(frmOptions.List1.hwnd, 0, X$ & vbNullString) = -1 Then
            frmOptions.List1.AddItem X$
        End If
    End If
Next i%
Unload Me

End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Command7_Click()
If List1.ListIndex > -1 And List1.ListIndex < List1.ListCount Then
For i% = 0 To UBound(cValid_eMails)
If cValid_eMails(i%).index = List1.ItemData(List1.ListIndex) Then
msgNote_tmp = cValid_eMails(i%).msgSource
frmViewMsg.Show 1, Me
Exit For
Else
msgNote_tmp = ""
End If
Next i%
End If


End Sub

Private Sub Command8_Click()
List1.Clear

End Sub

Private Sub Command9_Click()
If List1.ListIndex > -1 And List1.ListIndex < List1.ListCount Then
List1.RemoveItem List1.ListIndex
End If

End Sub

Private Sub Form_Load()
List2.AddItem "www.Planet-Source-Code.com"
List2.AddItem "Learn a little more about Visual Basic World"

End Sub


