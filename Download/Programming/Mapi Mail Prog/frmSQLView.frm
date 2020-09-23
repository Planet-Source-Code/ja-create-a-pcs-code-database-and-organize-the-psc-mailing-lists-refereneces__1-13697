VERSION 5.00
Begin VB.Form frmSQLView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View SQL"
   ClientHeight    =   4080
   ClientLeft      =   2220
   ClientTop       =   930
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Test"
      Height          =   390
      Left            =   4665
      TabIndex        =   3
      Top             =   3570
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Co&py"
      Height          =   390
      Left            =   5760
      TabIndex        =   2
      Top             =   3555
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   390
      Left            =   6840
      TabIndex        =   1
      Top             =   3570
      Width           =   945
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   45
      Width           =   7950
   End
End
Attribute VB_Name = "frmSQLView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text

End Sub

Private Sub Command3_Click()
Dim xSQL As String
xSQL = Text1.Text

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


Private Sub Form_Load()
If frmFilters.SQLtoVIEW = "SQLERROR" Then
Text1.Text = ""
MsgBox "Error in SQL"
Else
Text1.Text = frmFilters.SQLtoVIEW
End If



End Sub


