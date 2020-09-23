VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit entry"
   ClientHeight    =   5370
   ClientLeft      =   3960
   ClientTop       =   1920
   ClientWidth     =   6165
   Icon            =   "frmEditEntry.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh  lists"
      Height          =   885
      Left            =   4110
      Picture         =   "frmEditEntry.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1410
      Width           =   1005
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   5310
      TabIndex        =   28
      Top             =   1425
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   4905
      TabIndex        =   14
      Top             =   4905
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   3570
      TabIndex        =   13
      Top             =   4905
      Width           =   1200
   End
   Begin VB.CheckBox Check2 
      Height          =   210
      Left            =   1455
      TabIndex        =   11
      Top             =   4037
      Width           =   210
   End
   Begin VB.CheckBox Check1 
      Height          =   210
      Left            =   1455
      TabIndex        =   10
      Top             =   3790
      Width           =   210
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1455
      MaxLength       =   250
      TabIndex        =   9
      Top             =   3468
      Width           =   720
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1455
      MaxLength       =   250
      TabIndex        =   8
      Top             =   3146
      Width           =   720
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1470
      TabIndex        =   6
      Top             =   2442
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   36823
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1470
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2090
      Width           =   2550
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1470
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1738
      Width           =   2550
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1470
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1386
      Width           =   2550
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1470
      MaxLength       =   250
      TabIndex        =   2
      Top             =   1064
      Width           =   4560
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   1470
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4290
      Width           =   4560
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   1470
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   532
      Width           =   4560
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1470
      MaxLength       =   250
      TabIndex        =   0
      Top             =   210
      Width           =   4560
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   1470
      TabIndex        =   7
      Top             =   2794
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24379393
      CurrentDate     =   36823
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Note:"
      Height          =   225
      Index           =   12
      Left            =   90
      TabIndex        =   27
      Top             =   4290
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Checked:"
      Height          =   225
      Index           =   11
      Left            =   90
      TabIndex        =   26
      Top             =   4037
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Downloaded:"
      Height          =   225
      Index           =   10
      Left            =   90
      TabIndex        =   25
      Top             =   3790
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Value:"
      Height          =   225
      Index           =   9
      Left            =   90
      TabIndex        =   24
      Top             =   3468
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Accessed:"
      Height          =   225
      Index           =   8
      Left            =   90
      TabIndex        =   23
      Top             =   3146
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Listdate:"
      Height          =   225
      Index           =   7
      Left            =   90
      TabIndex        =   22
      Top             =   2794
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date submitted:"
      Height          =   225
      Index           =   6
      Left            =   90
      TabIndex        =   21
      Top             =   2442
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Compatibility:"
      Height          =   225
      Index           =   5
      Left            =   90
      TabIndex        =   20
      Top             =   2090
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Level:"
      Height          =   225
      Index           =   4
      Left            =   90
      TabIndex        =   19
      Top             =   1738
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Category:"
      Height          =   225
      Index           =   3
      Left            =   90
      TabIndex        =   18
      Top             =   1386
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Location:"
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   17
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   225
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   532
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Code Title:"
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F
Private Const LB_ERR = (-1)
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Sub EmptyTable(tb As DAO.Recordset)
On Error Resume Next
tb.MoveFirst
Do While Not (tb.EOF And tb.BOF)
tb.MoveFirst
tb.Delete
Loop




End Sub
Private Function FindStringExact(listHwnd As Long, ByVal IndexStart As Long, ByVal srcString As String) As Long
ret& = SendMessage(listHwnd, LB_FINDSTRING, IndexStart, ByVal srcString)
If ret& = LB_ERR Then
FindStringExact = -1
Else
FindStringExact = ret&
End If

End Function

Function GetComboIndex(d As String, cmb As ComboBox) As Integer
For i% = 0 To cmb.ListCount - 1
If cmb.List(i%) = d Then
GetComboIndex = i%
Exit Function
End If
Next i%

End Function

Private Sub Init()
Dim t As Database
Dim S As Recordset
Set t = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set S = t.OpenRecordset("LevelNames", dbOpenSnapshot)
Do While Not S.EOF
Combo2.AddItem S.Fields("Levelname").Value & vbNullString
S.MoveNext
Loop
S.Close
Set S = t.OpenRecordset("CategoryNames", dbOpenSnapshot)
Do While Not S.EOF
Combo1.AddItem S.Fields("Categoryname").Value & vbNullString
S.MoveNext
Loop
S.Close
Set S = t.OpenRecordset("CompatibilityNames", dbOpenSnapshot)
Do While Not S.EOF
Combo3.AddItem S.Fields("Compatibilityname").Value & vbNullString
S.MoveNext
Loop
S.Close
t.Close
ct& = SendMessage(Combo1.hwnd, CB_GETDROPPEDWIDTH, 0, 0)
ret& = SendMessage(Combo1.hwnd, CB_SETDROPPEDWIDTH, 1.5 * ct&, 0)
ret& = SendMessage(Combo2.hwnd, CB_SETDROPPEDWIDTH, 1 * ct&, 0)
ret& = SendMessage(Combo3.hwnd, CB_SETDROPPEDWIDTH, 2 * ct&, 0)


End Sub

Sub Refreshcombos()
MousePointer = 11
'/////////////////////////////////////////////////////////
'Category
Combo1.Visible = False
Refresh
List1.Clear
Dim t As Database
Dim S As Recordset
Set t = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set S = t.OpenRecordset("PSC1", dbOpenSnapshot)
Do While Not S.EOF
    If FindStringExact(List1.hwnd, 0, S.Fields("Category").Value & vbNullString) = -1 Then
    List1.AddItem S.Fields("Category").Value & vbNullString
    End If
    S.MoveNext
    
Loop
S.Close
Combo1.Clear
Set S = t.OpenRecordset("CategoryNames")
EmptyTable S
For i% = 0 To List1.ListCount - 1
S.AddNew
S.Fields("CategoryName").Value = List1.List(i%)
Combo1.AddItem List1.List(i%)
S.Update
Next i%
S.Close
Combo1.Visible = True
Refresh
'//////////////////////////////////////////////////////////////
'Level
Combo2.Visible = False
Refresh
List1.Clear
Set S = t.OpenRecordset("PSC1", dbOpenSnapshot)
Do While Not S.EOF
    If FindStringExact(List1.hwnd, 0, UCase(S.Fields("Level").Value & vbNullString)) = -1 Then
    List1.AddItem S.Fields("level").Value & vbNullString
    End If
    S.MoveNext
    
Loop
S.Close

Combo2.Clear
Set S = t.OpenRecordset("LevelNames")
EmptyTable S
For i% = 0 To List1.ListCount - 1
S.AddNew
S.Fields("LevelName").Value = List1.List(i%)
Combo2.AddItem List1.List(i%)
S.Update
Next i%
S.Close
Combo2.Visible = True
Refresh
'///////////////////////////////////////////////////////////////////
'Compatibility
Combo3.Visible = False
Refresh
List1.Clear
Set S = t.OpenRecordset("PSC1", dbOpenSnapshot)
Do While Not S.EOF
    If FindStringExact(List1.hwnd, 0, UCase(S.Fields("Compatibility").Value & vbNullString)) = -1 Then
    List1.AddItem S.Fields("Compatibility").Value & vbNullString
    End If
    S.MoveNext
    
Loop
S.Close
Combo3.Clear
Set S = t.OpenRecordset("CompatibilityNames")
EmptyTable S
For i% = 0 To List1.ListCount - 1
S.AddNew
S.Fields("CompatibilityName").Value = List1.List(i%)
Combo3.AddItem List1.List(i%)
S.Update
Next i%
S.Close
t.Close
List1.Clear
Combo3.Visible = True
MousePointer = 0
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Refresh

End Sub

Private Sub Check1_Click()
Command1.Enabled = True
End Sub

Private Sub Check2_Click()
Command1.Enabled = True
End Sub


Private Sub Combo1_Change()
Command1.Enabled = True
End Sub

Private Sub Combo1_Click()
Command1.Enabled = True
End Sub


Private Sub Combo2_Click()
Command1.Enabled = True
End Sub


Private Sub Combo3_Click()
Command1.Enabled = True
End Sub


Private Sub Command1_Click()
Dim cEntry As DAO.Fields
frmCodeEntries.Data1.Recordset.Edit
Set cEntry = frmCodeEntries.Data1.Recordset.Fields
cEntry("Title").Value = Text1(0).Text
cEntry("[Description]").Value = Text1(1).Text
cEntry("HTTP").Value = Text1(3).Text
cEntry("Category").Value = Combo1.Text
cEntry("Level").Value = Combo2.Text
cEntry("Compatibility").Value = Combo3.Text
cEntry("[DateSumbitted]").Value = DTPicker1.Value
cEntry("[Listdate]").Value = DTPicker2.Value
cEntry("Accessed").Value = Text1(4).Text
cEntry("[PersoanalValue]").Value = Text1(5).Text
cEntry("[PersonalNote]").Value = Text1(2).Text
cEntry("[HasBeendownloaded]").Value = CBool(check1.Value)
cEntry("[HasBeenChecked]").Value = CBool(Check2.Value)
frmCodeEntries.Data1.Recordset.Update
frmCodeEntries.Data1.UpdateControls
Command1.Enabled = False
Unload Me
End Sub

Private Sub Command2_Click()
        
Unload Me

End Sub

Private Sub Command3_Click()
x1$ = Combo1.Text
x2$ = Combo2.Text
x3$ = Combo3.Text
Refreshcombos
Combo1.ListIndex = GetComboIndex(x1$, Combo1)
Combo2.ListIndex = GetComboIndex(x2$, Combo2)
Combo3.ListIndex = GetComboIndex(x3$, Combo3)

End Sub

Private Sub DTPicker1_Change()
Command1.Enabled = True
End Sub


Private Sub DTPicker2_Change()
Command1.Enabled = True
End Sub


Private Sub Form_Load()
Init
Dim cEntry As DAO.Fields
Set cEntry = frmCodeEntries.Data1.Recordset.Fields
Text1(0).Text = cEntry("Title").Value & vbNullString
Text1(1).Text = cEntry("[Description]").Value & vbNullString
Text1(3).Text = cEntry("HTTP").Value & vbNullString
Combo1.ListIndex = GetComboIndex(cEntry("Category").Value, Combo1)
Combo2.ListIndex = GetComboIndex(cEntry("Level").Value, Combo2)
Combo3.ListIndex = GetComboIndex(cEntry("Compatibility").Value, Combo3)
If cEntry("[DateSumbitted]").Value < DTPicker1.MinDate Then
DTPicker1.Value = DTPicker1.MinDate
Else
DTPicker1.Value = cEntry("[DateSumbitted]").Value
End If
If cEntry("[Listdate]").Value < DTPicker2.MinDate Then
DTPicker2.Value = DTPicker2.MinDate
Else
DTPicker2.Value = cEntry("[Listdate]").Value
End If

Text1(4).Text = cEntry("Accessed").Value
Text1(5).Text = cEntry("[PersoanalValue]").Value
Text1(2).Text = cEntry("[PersonalNote]").Value & vbNullString
check1.Value = Abs(cEntry("[HasBeendownloaded]").Value)
Check2.Value = Abs(cEntry("[HasBeenChecked]").Value)

Command1.Enabled = False









End Sub

Private Sub Form_Unload(Cancel As Integer)
If Command1.Enabled = True Then
    ret& = MsgBox("Do you want to save the changes before closing ?", vbYesNoCancel + vbDefaultButton1 + vbQuestion, "Save changes")
        If ret& = vbYes Then
        Command1.Value = True
        Unload Me
        ElseIf ret& = vbCancel Then Cancel = True
        Exit Sub
        End If
End If


End Sub


Private Sub Text1_Change(index As Integer)
Command1.Enabled = True
End Sub

Private Sub Text1_GotFocus(index As Integer)
Text1(index).SelStart = 0
Text1(index).SelLength = Len(Text1(index).Text)
End Sub


