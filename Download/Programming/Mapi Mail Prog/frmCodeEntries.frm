VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCodeEntries 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PSC Code Entries"
   ClientHeight    =   6645
   ClientLeft      =   2340
   ClientTop       =   855
   ClientWidth     =   7965
   Icon            =   "frmCodeEntries.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7965
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   7965
      TabIndex        =   21
      Top             =   5805
      Width           =   7965
      Begin MSComDlg.CommonDialog CD1 
         Left            =   7830
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Filter"
         Height          =   450
         Left            =   4570
         TabIndex        =   45
         Top             =   0
         Width           =   1000
      End
      Begin VB.Data Data1 
         Caption         =   "PSC Code entries database"
         Connect         =   "Access"
         DatabaseName    =   "D:\Download\Programming\Mapi Mail Prog\PSC.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   75
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "PSC1"
         Top             =   465
         Width           =   7785
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Exit"
         Height          =   450
         Left            =   6810
         TabIndex        =   27
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Options"
         Height          =   450
         Left            =   5690
         TabIndex        =   26
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Search"
         Height          =   450
         Left            =   3450
         TabIndex        =   25
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   450
         Left            =   2330
         TabIndex        =   24
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Edit"
         Height          =   450
         Left            =   1210
         TabIndex        =   23
         Top             =   0
         Width           =   1000
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add"
         Height          =   450
         Left            =   90
         TabIndex        =   22
         Top             =   0
         Width           =   1000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record 1/1"
      Height          =   5625
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   7755
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   6705
         MaskColor       =   &H00FF00FF&
         Picture         =   "frmCodeEntries.frx":06C2
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2265
         UseMaskColor    =   -1  'True
         Width           =   360
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   2745
         Left            =   150
         ScaleHeight     =   2685
         ScaleWidth      =   7410
         TabIndex        =   8
         Top             =   2760
         Width           =   7470
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Height          =   270
            Left            =   7050
            Picture         =   "frmCodeEntries.frx":07C4
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Locate"
            Top             =   1695
            Width           =   315
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00000000&
            DataField       =   "PersonalNote"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   480
            Left            =   570
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "txtNote"
            Top             =   2085
            Width           =   6600
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   5040
            Picture         =   "frmCodeEntries.frx":08C6
            Top             =   1380
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   5040
            Picture         =   "frmCodeEntries.frx":09C8
            Top             =   1110
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Note"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   15
            Left            =   75
            TabIndex        =   41
            Top             =   2205
            Width           =   600
         End
         Begin VB.Label lblLocalDir 
            BackStyle       =   0  'Transparent
            Caption         =   "lblLocalDir"
            DataField       =   "LocalDirectory"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   5040
            TabIndex        =   39
            Top             =   1725
            Width           =   1995
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Local Dir"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   14
            Left            =   3585
            TabIndex        =   38
            Top             =   1725
            Width           =   900
         End
         Begin VB.Label lblChecked 
            BackStyle       =   0  'Transparent
            Caption         =   "lblChecked"
            DataField       =   "HasBeenChecked"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   5400
            TabIndex        =   37
            Top             =   1410
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Checked"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   13
            Left            =   3585
            TabIndex        =   36
            Top             =   1410
            Width           =   900
         End
         Begin VB.Label lblDownloaded 
            BackStyle       =   0  'Transparent
            Caption         =   "lblDownloaded"
            DataField       =   "HasBeenDownloaded"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   5400
            TabIndex        =   35
            Top             =   1110
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Downloaded"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   12
            Left            =   3585
            TabIndex        =   34
            Top             =   1110
            Width           =   900
         End
         Begin VB.Label lblValue 
            BackStyle       =   0  'Transparent
            Caption         =   "lblValue"
            DataField       =   "PersoanalValue"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   5040
            TabIndex        =   33
            Top             =   750
            Width           =   2000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   11
            Left            =   3585
            TabIndex        =   32
            Top             =   750
            Width           =   900
         End
         Begin VB.Label lblmsgID 
            BackStyle       =   0  'Transparent
            Caption         =   "lblmsgID"
            DataField       =   "msgID"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   5040
            TabIndex        =   31
            Top             =   420
            Width           =   2000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "msgID"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   10
            Left            =   3585
            TabIndex        =   30
            Top             =   420
            Width           =   900
         End
         Begin VB.Label lblCodeID 
            BackStyle       =   0  'Transparent
            Caption         =   "lblCodeID"
            DataField       =   "IndexID"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   5040
            TabIndex        =   29
            Top             =   90
            Width           =   2000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Code ID"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   9
            Left            =   3585
            TabIndex        =   28
            Top             =   90
            Width           =   900
         End
         Begin VB.Label lblAcessed 
            BackStyle       =   0  'Transparent
            Caption         =   "lblAcessed"
            DataField       =   "Accessed"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1485
            TabIndex        =   20
            Top             =   1755
            Width           =   2000
         End
         Begin VB.Label lblListDate 
            BackStyle       =   0  'Transparent
            Caption         =   "lblListDate"
            DataField       =   "ListDate"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1485
            TabIndex        =   19
            Top             =   1455
            Width           =   2000
         End
         Begin VB.Label lblDateSubmitted 
            BackStyle       =   0  'Transparent
            Caption         =   "lblDateSubmitted"
            DataField       =   "DateSumbitted"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1485
            TabIndex        =   18
            Top             =   1095
            Width           =   2000
         End
         Begin VB.Label lblCompatibility 
            BackStyle       =   0  'Transparent
            Caption         =   "lblCompatibility"
            DataField       =   "Compatibility"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1485
            TabIndex        =   17
            Top             =   750
            Width           =   2000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   30
            TabIndex        =   16
            Top             =   90
            Width           =   900
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Level"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   30
            TabIndex        =   15
            Top             =   420
            Width           =   870
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Compatibility"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   4
            Left            =   30
            TabIndex        =   14
            Top             =   750
            Width           =   1080
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Submitted"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   5
            Left            =   30
            TabIndex        =   13
            Top             =   1095
            Width           =   1320
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Listdate"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   6
            Left            =   30
            TabIndex        =   12
            Top             =   1425
            Width           =   1245
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Accessed"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   7
            Left            =   30
            TabIndex        =   11
            Top             =   1755
            Width           =   1365
         End
         Begin VB.Label lblCategory 
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            DataField       =   "Category"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1485
            TabIndex        =   10
            Top             =   90
            Width           =   2000
         End
         Begin VB.Label lblLevel 
            BackStyle       =   0  'Transparent
            Caption         =   "lblLevel"
            DataField       =   "Level"
            DataSource      =   "Data1"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1485
            TabIndex        =   9
            Top             =   420
            Width           =   2000
         End
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00000000&
         DataField       =   "Title"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   480
         Left            =   150
         TabIndex        =   4
         Top             =   525
         Width           =   7470
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00000000&
         DataField       =   "Description"
         DataSource      =   "Data1"
         ForeColor       =   &H00C0FFC0&
         Height          =   690
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1275
         Width           =   7470
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Http"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Text            =   "http://www.sdfsdf.2wrwe/dsfsdfsdfs/sdfsdfsd/index.htm"
         Top             =   2265
         Width           =   6450
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   7110
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2265
         UseMaskColor    =   -1  'True
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Code of the Day"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1170
         TabIndex        =   43
         Top             =   150
         Visible         =   0   'False
         Width           =   4530
      End
      Begin VB.Label Label1 
         Caption         =   "Code Title"
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   285
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   210
         Index           =   1
         Left            =   225
         TabIndex        =   6
         Top             =   1020
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Location"
         Height          =   210
         Index           =   8
         Left            =   225
         TabIndex        =   5
         Top             =   2055
         Width           =   990
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "Filters"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenFilter 
         Caption         =   "Open filter"
      End
      Begin VB.Menu mnuCreateFilter 
         Caption         =   "Create new filter"
      End
      Begin VB.Menu mnuNoFilters 
         Caption         =   "No Filters"
      End
   End
End
Attribute VB_Name = "frmCodeEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FilterSQL As String

Public Sub RequeryDatabase(SQL As String)
Data1.DatabaseName = App.path & "\psc.mdb"
FilterSQL = SQL
Data1.RecordSource = FilterSQL
Data1.Refresh




End Sub


Private Sub Command1_Click()
If Trim(Text1.Text) <> "" Then
openURL Text1.Text
Else
MsgBox "No URL specified"
End If

End Sub

Private Sub Command10_Click()
Dim xCO As Double, yCo As Double
xCO = Picture2.Left + Command10.Left
yCo = Picture2.Top + Command10.Top + Command10.Height
PopupMenu mnuFilters, , xCO, yCo

End Sub

Private Sub Command2_Click()
Form1.Show 1, Me



End Sub

Private Sub Command3_Click()
frmEditEntry.Show 1, Me

End Sub

Private Sub Command4_Click()
If MsgBox("Are you sure that you want to delete the current code entry ?", vbYesNo + vbQuestion + vbDefaultButton1, "Delete PSC code entry") = vbNo Then Exit Sub

If Data1.Recordset.AbsolutePosition = Data1.Recordset.RecordCount Then
Data1.Recordset.Delete
Data1.Recordset.MoveLast
ElseIf Data1.Recordset.AbsolutePosition = 0 Then
Data1.Recordset.Delete
Data1.Recordset.MoveFirst
Else
Data1.Recordset.Delete
Data1.Recordset.MovePrevious
End If

End Sub

Private Sub Command5_Click()
frmSearch.Show

End Sub

Private Sub Command6_Click()
frmOptions.Show 1, Me

End Sub

Private Sub Command7_Click()
Unload Me
End

End Sub

Private Sub Command8_Click()
Dim strResFolder As String
ChDir "d:\download\programming"
strResFolder = BrowseForFolder(hwnd, "Please select a folder.")

If strResFolder <> "" Then
lblLocalDir.Caption = strResFolder
Data1.Recordset.Edit
Data1.Recordset.Update
End If


End Sub

Public Sub Command9_Click()
If Len(Text1.Text) > 0 Then
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Clipboard.Clear
Clipboard.SetText Text1.Text
Text1.SetFocus
End If

End Sub

Private Sub Data1_Reposition()
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
Frame1.Caption = "Record " & Data1.Recordset.AbsolutePosition + 1 & " from " & Data1.Recordset.RecordCount
Command8.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Label2.Visible = Data1.Recordset.Fields("[CodeOfTheDay]").Value
Else
Frame1.Caption = "No Records found"
Command8.Enabled = False
Command3.Enabled = 0
Command4.Enabled = 0
Label2.Visible = False
End If


End Sub

Private Sub Form_Initialize()
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
Data1.Recordset.MoveLast
End If

End Sub

Private Sub Form_Load()
Dim cIc As Integer
cIc = GetBrowseCursorResID

Select Case cIc
    Case 103
    Command1.Caption = ""
    Command1.Picture = LoadResPicture(cIc + 4, vbResBitmap)
    Command1.Width = 360
    Case 104
    Command1.Caption = ""
    Command1.Picture = LoadResPicture(cIc + 4, vbResBitmap)
    Command1.Width = 360
    Case 105
    Command1.Width = 555
    Command1.Caption = "Open"
End Select
LoadStartUpSettings
Data1.DatabaseName = App.path & "\psc.mdb"
If LoadSQLatStartup = True Then
FilterSQL = StartUpSQL
Caption = "PSC Code entries (Filters=ON)"
Else
FilterSQL = "Select PSC1.* FROM PSC1 order by IndexID ASC,datesumbitted DESC,Accessed DESC"
Caption = "PSC Code entries (Filters=OFF)"
End If

Data1.RecordSource = FilterSQL
Data1.Refresh


End Sub


Private Sub Form_Unload(Cancel As Integer)
End

End Sub


Private Sub lblAcessed_Change()
lblAcessed.ToolTipText = lblAcessed.Caption

End Sub

Private Sub lblCategory_Change()
lblCategory.ToolTipText = lblCategory.Caption
End Sub

Private Sub lblChecked_Change()
Image2.Visible = lblChecked.Caption = "-1"

End Sub

Private Sub lblCodeID_Change()
lblCodeID.ToolTipText = lblCodeID

End Sub

Private Sub lblCompatibility_Change()
lblCompatibility.ToolTipText = lblCompatibility.Caption

End Sub

Private Sub lblDateSubmitted_Change()
lblDateSubmitted.ToolTipText = lblDateSubmitted.Caption

End Sub

Private Sub lblDownloaded_Change()
Image1.Visible = lblDownloaded.Caption = "-1"


End Sub

Private Sub lblLevel_Change()
lblLevel.ToolTipText = lblLevel.Caption
End Sub

Private Sub lblListDate_Change()
lblListDate.ToolTipText = lblListDate.Caption
If Label2.Visible Then Label2 = "Code of the Day (" & lblListDate.Caption & ")"

End Sub

Private Sub lblLocalDir_Change()
lblLocalDir.ToolTipText = lblLocalDir
End Sub

Private Sub lblmsgID_Change()
lblmsgID.ToolTipText = lblmsgID

End Sub

Private Sub lblValue_Change()
lblValue.ToolTipText = lblValue
End Sub



Private Sub mnuCreateFilter_Click()
frmFilters.Show 1, Me

End Sub


Private Sub mnuNoFilters_Click()
MousePointer = 11
FilterSQL = "Select PSC1.* FROM PSC1 order by IndexID ASC,datesumbitted DESC,Accessed DESC"
RequeryDatabase FilterSQL
Caption = "PSC Code entries (Filters=OFF)"
MousePointer = 0
End Sub

Private Sub mnuOpenFilter_Click()
cd1.DialogTitle = "Open saved filter"
cd1.InitDir = App.path
cd1.Flags = cdlOFNPathMustExist + cdlOFNFileMustExist + cdlOFNHideReadOnly
cd1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
cd1.FilterIndex = 0
cd1.DefaultExt = "txt"
cd1.CancelError = True
On Error GoTo cderr3
cd1.ShowOpen
MousePointer = 11
If cd1.FileName <> "" Then
r% = FreeFile
Dim X As String * 1
Dim aSQl As String
Open cd1.FileName For Binary As #r% Len = 1
Do While Not EOF(r%)
Get #r%, , X
aSQl = aSQl & X
Loop
If TestSQL(aSQl) Then
RequeryDatabase aSQl
FilterSQL = aSQl
Caption = "PSC Code entries (Filters=ON)"
Else
MsgBox "Not a valid SQL string"
End If
Close #r%
End If
MousePointer = 0

Exit Sub

cderr3:
Close
MousePointer = 0
Exit Sub
Resume Next



End Sub


