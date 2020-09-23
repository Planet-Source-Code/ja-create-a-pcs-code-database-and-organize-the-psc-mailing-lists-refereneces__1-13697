VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFilters 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filters"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmFilters.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7170
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command17 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7905
      TabIndex        =   53
      Top             =   5700
      Width           =   930
   End
   Begin VB.CommandButton Command16 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5985
      TabIndex        =   52
      Top             =   5700
      Width           =   930
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Test SQL"
      Height          =   375
      Left            =   7905
      TabIndex        =   51
      Top             =   5295
      Width           =   930
   End
   Begin VB.CommandButton Command14 
      Caption         =   "&Save SQL"
      Height          =   375
      Left            =   6960
      TabIndex        =   50
      Top             =   5475
      Width           =   930
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&View SQL"
      Height          =   375
      Left            =   5985
      TabIndex        =   49
      Top             =   5295
      Width           =   930
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      Height          =   195
      Left            =   1095
      TabIndex        =   48
      Top             =   5250
      Width           =   210
   End
   Begin VB.Frame Frame6 
      Caption         =   "Date filters"
      Enabled         =   0   'False
      Height          =   840
      Left            =   150
      TabIndex        =   43
      Top             =   5235
      Width           =   5700
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   825
         TabIndex        =   44
         Top             =   405
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   36824
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   3810
         TabIndex        =   46
         Top             =   405
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   36824
      End
      Begin VB.Label Label6 
         Caption         =   "and listdate<="
         Height          =   240
         Left            =   2670
         TabIndex        =   47
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Listdate>="
         Height          =   240
         Left            =   75
         TabIndex        =   45
         Top             =   435
         Width           =   780
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Include if they"
      Height          =   1590
      Left            =   5970
      TabIndex        =   38
      Top             =   3540
      Width           =   2850
      Begin VB.CheckBox Check3 
         Caption         =   "have been checked"
         Height          =   330
         Left            =   135
         TabIndex        =   41
         Top             =   830
         Width           =   1875
      End
      Begin VB.CheckBox Check4 
         Caption         =   "don't have been checked"
         Height          =   330
         Left            =   135
         TabIndex        =   42
         Top             =   1125
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "don't have been downloaded"
         Height          =   330
         Left            =   135
         TabIndex        =   40
         Top             =   535
         Width           =   2460
      End
      Begin VB.CheckBox Check1 
         Caption         =   "have been downloaded"
         Height          =   330
         Left            =   135
         TabIndex        =   39
         Top             =   240
         Width           =   2370
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sort by"
      Height          =   3285
      Left            =   5970
      TabIndex        =   21
      Top             =   135
      Width           =   2835
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   15
         ScaleHeight     =   645
         ScaleWidth      =   2700
         TabIndex        =   34
         Top             =   2520
         Width           =   2700
         Begin VB.OptionButton Option8 
            Caption         =   "Descending"
            Height          =   255
            Left            =   1230
            TabIndex        =   36
            Top             =   300
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Ascending"
            Height          =   255
            Left            =   1230
            TabIndex        =   35
            Top             =   0
            Width           =   1410
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Code ID"
            Height          =   210
            Left            =   0
            TabIndex        =   37
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   15
         ScaleHeight     =   645
         ScaleWidth      =   2730
         TabIndex        =   30
         Top             =   1755
         Width           =   2730
         Begin VB.OptionButton Option6 
            Caption         =   "Ascending"
            Height          =   255
            Left            =   1230
            TabIndex        =   32
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Descending"
            Height          =   255
            Left            =   1230
            TabIndex        =   31
            Top             =   300
            Width           =   1380
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Accessed"
            Height          =   210
            Left            =   0
            TabIndex        =   33
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   15
         ScaleHeight     =   645
         ScaleWidth      =   2760
         TabIndex        =   26
         Top             =   990
         Width           =   2760
         Begin VB.OptionButton Option4 
            Caption         =   "Ascending"
            Height          =   255
            Left            =   1230
            TabIndex        =   28
            Top             =   0
            Width           =   1485
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Descending"
            Height          =   255
            Left            =   1230
            TabIndex        =   27
            Top             =   300
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Listdate"
            Height          =   210
            Left            =   0
            TabIndex        =   29
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   15
         ScaleHeight     =   645
         ScaleWidth      =   2715
         TabIndex        =   22
         Top             =   225
         Width           =   2715
         Begin VB.OptionButton Option2 
            Caption         =   "Descending"
            Height          =   255
            Left            =   1230
            TabIndex        =   25
            Top             =   300
            Value           =   -1  'True
            Width           =   1440
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ascending"
            Height          =   255
            Left            =   1230
            TabIndex        =   24
            Top             =   0
            Width           =   1440
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date submitted"
            Height          =   210
            Left            =   0
            TabIndex        =   23
            Top             =   180
            Width           =   1155
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Categories to include"
      Height          =   1590
      Left            =   165
      TabIndex        =   14
      Top             =   3540
      Width           =   5700
      Begin VB.ListBox List6 
         Height          =   1230
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2370
      End
      Begin VB.ListBox List5 
         Height          =   1230
         Left            =   3210
         TabIndex        =   19
         Top             =   240
         Width           =   2370
      End
      Begin VB.CommandButton Command12 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   18
         Top             =   255
         Width           =   450
      End
      Begin VB.CommandButton Command11 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   17
         Top             =   575
         Width           =   450
      End
      Begin VB.CommandButton Command10 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   16
         Top             =   895
         Width           =   450
      End
      Begin VB.CommandButton Command9 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   15
         Top             =   1215
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Levels to include"
      Height          =   1590
      Left            =   165
      TabIndex        =   7
      Top             =   1837
      Width           =   5700
      Begin VB.ListBox List4 
         Height          =   1230
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2370
      End
      Begin VB.ListBox List3 
         Height          =   1230
         Left            =   3210
         TabIndex        =   12
         Top             =   240
         Width           =   2370
      End
      Begin VB.CommandButton Command8 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   11
         Top             =   255
         Width           =   450
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   10
         Top             =   575
         Width           =   450
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   9
         Top             =   895
         Width           =   450
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   8
         Top             =   1215
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Categories to include"
      Height          =   1590
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   5700
      Begin VB.CommandButton Command4 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   6
         Top             =   1215
         Width           =   450
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   5
         Top             =   895
         Width           =   450
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   4
         Top             =   575
         Width           =   450
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   3
         Top             =   255
         Width           =   450
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Left            =   3210
         TabIndex        =   2
         Top             =   240
         Width           =   2370
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2370
      End
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LB_ERR = (-1)
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Declare Function LBSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public SQLtoVIEW As String
Function TestSQL() As Boolean
'Checks if a SQL string is valid

Dim xSQL As String
xSQL = CreateSQL
If xSQL = "SQLERROR" Then
TestSQL = False
Exit Function
End If

On Error GoTo errSQL1
Dim testDB As Database
Dim testSet As Recordset
Set testDB = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set testSet = testDB.OpenRecordset(xSQL, dbOpenDynaset)
testSet.Close
testDB.Close
Set testDB = Nothing
Set testSet = Nothing
TestSQL = True
Exit Function


errSQL1:
TestSQL = False
Exit Function
Resume Next


End Function


Function CreateSQL() As String

Dim Bound1 As Integer, Bound2 As Integer, Bound3 As Integer
Dim cCategory() As String
Dim cLevel() As String
Dim cCompatibility() As String
Dim cSQL As String
Dim mSQL As String
Dim lSQL As String
Dim jSQL As String
Dim DateSQL As String
Dim x1 As String, x2 As String
Dim checkSQL As String
Dim orderSQL As String


If List2.ListCount = 0 And List3.ListCount = 0 And List5.ListCount = 0 Then
MsgBox "You must select at least one Level,Category or Compatibility from the lists.", vbCritical, "SQL error"
CreateSQL = "SQLERROR"
Exit Function
End If

Bound1 = List2.ListCount - 1
Bound2 = List3.ListCount - 1
Bound3 = List5.ListCount - 1

If Bound1 > -1 Then ReDim cCategory(0 To Bound1)
If Bound2 > -1 Then ReDim cLevel(0 To Bound2)
If Bound3 > -1 Then ReDim cCompatibility(0 To Bound3)

cSQL = "Select PCS.* FROM PSC "
'''''''''''''''''
'Category
'''''''''''''''''
For i% = 0 To Bound1
    cCategory(i%) = "[Category]='" & List2.List(i%) & "'"
Next i%
mSQL = ""
For i% = 0 To Bound1
mSQL = mSQL & cCategory(i%) & " or "
Next i%
''''''
'create first (and)
'''''''
If Len(mSQL) > 0 Then
mSQL = "(" & Left$(mSQL, Len(mSQL) - 4) & ")"
End If
''''''''''''''''''''
'Level
''''''''''''''''''''
For i% = 0 To Bound2
    cLevel(i%) = "[Level]='" & List3.List(i%) & "'"
Next i%
lSQL = ""
For i% = 0 To Bound2
lSQL = lSQL & cLevel(i%) & " or "
Next i%
''''''
'create second (and)
'''''''
If Len(lSQL) > 0 Then
lSQL = "(" & Left$(lSQL, Len(lSQL) - 4) & ")"
End If
''''''''''''''''''''
'Compatibility
''''''''''''''''''''
For i% = 0 To Bound3
    cCompatibility(i%) = "[Compatibility]='" & List5.List(i%) & "'"
Next i%
jSQL = ""
For i% = 0 To Bound3
jSQL = jSQL & cCompatibility(i%) & " or "
Next i%
''''''
'create third (and)
'''''''
If Len(jSQL) > 0 Then
jSQL = "(" & Left$(jSQL, Len(jSQL) - 4) & ")"
End If
''''''''''
'Create 4th (and)
''''''''''
If Check5.Value = vbChecked Then
    If DTPicker1.Value > DTPicker2.Value Then
    MsgBox "The 'Before date' value is larger from the 'After date' value. Please change the values and retry.", vbCritical, "Error in SQL"
    CreateSQL = "SQLERROR"
    Exit Function
    Else
    DateSQL = "(Listdate>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and Listdate<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "#)"
    End If
End If

''''''''''
'Create 5th (and)
''''''''''
If check1.Value = vbUnchecked And Check2.Value = vbUnchecked Then
x1 = "([HasbeenDownloaded]=TRUE or [HasbeenDownloaded]=false)"
Else
    If check1.Value = vbChecked Then
    x1 = "([HasbeenDownloaded]=TRUE)"
    Else
    x1 = "([HasbeenDownloaded]=false)"
    End If
End If

If Check3.Value = vbUnchecked And Check4.Value = vbUnchecked Then
x2 = "([Hasbeenchecked]=TRUE or [Hasbeenchecked]=false)"
Else
    If Check3.Value = vbChecked Then
    x2 = "([Hasbeenchecked]=TRUE)"
    Else
    x2 = "([Hasbeenchecked]=False)"
    End If
    
End If

checkSQL = "(" & x1 & " and " & x2 & ")"

cSQL = "Select PSC1.* FROM PSC1 WHERE(" & mSQL
If lSQL <> "" Then
    If mSQL <> "" Then
    cSQL = cSQL & " and " & lSQL
    Else
    cSQL = cSQL & lSQL
    End If
End If

If jSQL <> "" Then
    If mSQL <> "" Or lSQL <> "" Then
    cSQL = cSQL & " and " & jSQL
    Else
    cSQL = cSQL & jSQL
    End If
End If

If DateSQL <> "" Then
    cSQL = cSQL & " and " & DateSQL
End If
cSQL = cSQL & " and (" & checkSQL & ")"

    


cSQL = cSQL & ")"
If Option1.Value Then
cSQL = cSQL & " order by Datesumbitted ASC, "
Else
cSQL = cSQL & " order by Datesumbitted DESC, "
End If

If Option3.Value Then
cSQL = cSQL & "listdate ASC, "
Else
cSQL = cSQL & "listdate DESC, "
End If

If Option5.Value Then
cSQL = cSQL & "accessed ASC, "
Else
cSQL = cSQL & "accessed DESC, "
End If

If Option7.Value Then
cSQL = cSQL & "IndexID ASC"
Else
cSQL = cSQL & "IndexID DESC"
End If



CreateSQL = cSQL


Exit Function










End Function


Private Function FindStringExact(ByVal lbHWND, ByVal IndexStart As Long, ByVal srcString As String) As Boolean
ret& = LBSendMessage(lbHWND, &H1A2, IndexStart, ByVal srcString)
If ret& = LB_ERR Then
FindStringExact = False
Else
FindStringExact = True
End If

End Function

Sub LoadLists()
Dim tDB As Database
Dim tSet As Recordset
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear

Set tDB = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set tSet = tDB.OpenRecordset("CategoryNames", dbOpenSnapshot)
If Not (tSet.EOF And tSet.BOF) Then
    Do While Not tSet.EOF
    List1.AddItem tSet.Fields(0).Value
    List2.AddItem tSet.Fields(0).Value
    tSet.MoveNext
    Loop
End If
tSet.Close
Set tSet = Nothing

Set tSet = tDB.OpenRecordset("LevelNames", dbOpenSnapshot)
If Not (tSet.EOF And tSet.BOF) Then
    Do While Not tSet.EOF
    List3.AddItem tSet.Fields(0).Value
    List4.AddItem tSet.Fields(0).Value
    tSet.MoveNext
    Loop
End If
tSet.Close
Set tSet = Nothing

Set tSet = tDB.OpenRecordset("CompatibilityNames", dbOpenSnapshot)
If Not (tSet.EOF And tSet.BOF) Then
    Do While Not tSet.EOF
    List5.AddItem tSet.Fields(0).Value
    List6.AddItem tSet.Fields(0).Value
    tSet.MoveNext
    Loop
End If
tSet.Close
Set tSet = Nothing
tDB.Close
Set tDB = Nothing

If List1.ListCount > 0 Then
List1.ListIndex = 0
End If
If List2.ListCount > 0 Then
List2.ListIndex = 0
End If
If List3.ListCount > 0 Then
List3.ListIndex = 0
End If
If List4.ListCount > 0 Then
List4.ListIndex = 0
End If
If List5.ListCount > 0 Then
List5.ListIndex = 0
End If
If List6.ListCount > 0 Then
List6.ListIndex = 0
End If


End Sub


Private Sub Check1_Click()
If check1.Value = vbChecked Then Check2.Value = vbUnchecked

End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then check1.Value = vbUnchecked

End Sub


Private Sub Check3_Click()
If Check3.Value = vbChecked Then Check4.Value = vbUnchecked

End Sub


Private Sub Check4_Click()
If Check4.Value = vbChecked Then Check3.Value = vbUnchecked

End Sub


Private Sub Check5_Click()
Frame6.Enabled = CBool(Check5.Value)

End Sub


Private Sub Command1_Click()
x$ = List1.List(List1.ListIndex) & vbNullString

If FindStringExact(List2.hwnd, 0, x$) = False Then
List2.AddItem x$
End If
List2.ListIndex = List2.NewIndex

End Sub

Private Sub Command10_Click()
If List5.ListCount = 1 Then
List5.Clear
Exit Sub
End If

i% = List5.ListIndex
If List5.ListIndex > -1 Then
List5.RemoveItem i%
End If

If i% - 1 > -1 Then
    List5.ListIndex = i% - 1
ElseIf i% <> 1 Then
    If i% < List5.ListCount - 1 Then
        List5.ListIndex = i
    End If
Else
    List5.ListIndex = 0
End If
End Sub

Private Sub Command11_Click()
For i% = 0 To List6.ListCount - 1
x$ = List6.List(i%) & vbNullString
If FindStringExact(List5.hwnd, 0, x$) = False Then
List5.AddItem x$
End If
Next i%
List5.ListIndex = List5.NewIndex
End Sub

Private Sub Command12_Click()
x$ = List6.List(List6.ListIndex) & vbNullString

If FindStringExact(List5.hwnd, 0, x$) = False Then
List5.AddItem x$
End If
List5.ListIndex = List5.NewIndex
End Sub

Private Sub Command13_Click()
SQLtoVIEW = CreateSQL
frmSQLView.Show 1
End Sub

Private Sub Command14_Click()
Dim xSQL As String
xSQL = CreateSQL

CD1.DialogTitle = "Save SQL string"
CD1.InitDir = App.path
CD1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*)"
CD1.FilterIndex = 0
CD1.DefaultExt = "txt"
CD1.Flags = cdlOFNOverwritePrompt + cdlOFNExtensionDifferent + cdlOFNPathMustExist + cdlOFNHideReadOnly
CD1.CancelError = True
On Error GoTo cdErr
CD1.ShowSave
If CD1.FileName <> "" Then
r% = FreeFile
Open CD1.FileName For Output As #r%
Print #r%, Trim(xSQL)
Close #r%
End If
Exit Sub
cdErr:
MsgBox Error
Close
Exit Sub
Resume Next

End Sub

Private Sub Command15_Click()
Dim xSQL As String
xSQL = CreateSQL
If xSQL = "SQLERROR" Then
Exit Sub
End If

On Error GoTo errSQL
Dim testDB As Database
Dim textSet As Recordset
Set testDB = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set testSet = testDB.OpenRecordset(xSQL, dbOpenDynaset)
testSet.Close
testDB.Close
Set testDB = Nothing
Set testSet = Nothing
MsgBox "Success" & vbCrLf & "No errors found", vbInformation, "SQL test results"


Exit Sub
errSQL:
MsgBox "Error in SQL." & vbCrLf & "Error ID: " & Err & vbCrLf & Error, vbCritical, "SQL Test results"
Exit Sub
Resume Next

End Sub

Private Sub Command16_Click()
Dim xSQL As String
If TestSQL = True Then
xSQL = CreateSQL
frmCodeEntries.RequeryDatabase xSQL
frmCodeEntries.FilterSQL = xSQL
frmCodeEntries.Caption = "PSC Code entries (Filters=ON)"
Unload Me
Else
MsgBox "Not as valid SQL string"
End If

End Sub

Private Sub Command17_Click()
Unload Me

End Sub


Private Sub Command2_Click()
For i% = 0 To List1.ListCount - 1
x$ = List1.List(i%) & vbNullString
If FindStringExact(List2.hwnd, 0, x$) = False Then
List2.AddItem x$
End If
Next i%
List2.ListIndex = List2.NewIndex

End Sub

Private Sub Command3_Click()
If List2.ListCount = 1 Then
List2.Clear
Exit Sub
End If

i% = List2.ListIndex
If List2.ListIndex > -1 Then
List2.RemoveItem i%
End If

If i% - 1 > -1 Then
    List2.ListIndex = i% - 1
ElseIf i% <> 1 Then
    If i% < List2.ListCount - 1 Then
        List2.ListIndex = i
    End If
Else
    List2.ListIndex = 0
End If


End Sub

Private Sub Command4_Click()
List2.Clear

End Sub


Private Sub Command5_Click()
List3.Clear
End Sub


Private Sub Command6_Click()

If List3.ListCount = 1 Then
List3.Clear
Exit Sub
End If

i% = List3.ListIndex
If List3.ListIndex > -1 Then
List3.RemoveItem i%
End If

If i% - 1 > -1 Then
    List3.ListIndex = i% - 1
ElseIf i% <> 1 Then
    If i% < List3.ListCount - 1 Then
        List3.ListIndex = i
    End If
Else
    List3.ListIndex = 0
End If

End Sub

Private Sub Command7_Click()
For i% = 0 To List4.ListCount - 1
x$ = List4.List(i%) & vbNullString
If FindStringExact(List3.hwnd, 0, x$) = False Then
List3.AddItem x$
End If
Next i%
List3.ListIndex = List3.NewIndex

End Sub

Private Sub Command8_Click()
x$ = List4.List(List4.ListIndex) & vbNullString

If FindStringExact(List3.hwnd, 0, x$) = False Then
List3.AddItem x$
End If
List3.ListIndex = List3.NewIndex
End Sub


Private Sub Command9_Click()
List5.Clear
End Sub

Private Sub Form_Load()
LoadLists

End Sub

Private Sub List1_DblClick()
Command1_Click

End Sub


Private Sub List2_DblClick()
Command3_Click

End Sub


Private Sub List3_DblClick()
Command6_Click

End Sub


Private Sub List4_DblClick()
Command8_Click

End Sub


Private Sub List5_DblClick()
Command9_Click
End Sub


Private Sub List6_DblClick()
Command12_Click

End Sub


