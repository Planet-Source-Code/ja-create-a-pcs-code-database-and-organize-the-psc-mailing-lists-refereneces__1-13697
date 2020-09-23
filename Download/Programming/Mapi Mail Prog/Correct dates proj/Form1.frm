VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   135
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ListdateYeartoChange As Long
Public DatesubmittedYeartoChange As Long





Private Sub Command1_Click()
Dim tdb As DAO.Database
Dim tset As DAO.Recordset
Dim dDate As Date
Dim sDate As Date
Dim counter As Integer
Dim counterchanged As Integer
Dim OldYearD As Long, OldYearS As Long

Set tdb = Workspaces(0).OpenDatabase("d:\download\programming\Mapi Mail Prog\PSC.mdb")
Set tset = tdb.OpenRecordset("PSC1")
Do While Not tset.EOF
    dDate = tset.Fields("Listdate").Value
    sDate = tset.Fields("DateSumbitted").Value
    OldYearD = Year(dDate)
    OldYearS = Year(sDate)
    If Year(dDate) < 1990 Or Year(sDate) < 1990 Then
    counter = counter + 1
    ListdateYeartoChange = Year(dDate)
    DatesubmittedYeartoChange = Year(sDate)
    If ListdateYeartoChange = 200 Then ListdateYeartoChange = 2000
    If DatesubmittedYeartoChange = 200 Then DatesubmittedYeartoChange = 2000
    frmChange.Show 1, Me
        If OldYearD <> ListdateYeartoChange Or OldYearS <> DatesubmittedYeartoChange Then
            dDate = CDate(CStr(Day(dDate)) & "/" & CStr(Month(dDate)) & "/" & CStr(ListdateYeartoChange))
            sDate = CDate(CStr(Day(sDate)) & "/" & CStr(Month(sDate)) & "/" & CStr(DatesubmittedYeartoChange))
            tset.Edit
            tset.Fields("Listdate").Value = dDate
            tset.Fields("DateSumbitted").Value = sDate
            tset.Update
            counterchanged = counterchanged + 1
        End If
    End If
    tset.MoveNext
    Me.Caption = counter
    Loop
    
tset.Close
tdb.Close
MsgBox "Found " & counter & "entries with invalid dates" & vbCrLf & "Corrected " & counterchanged & entries
Unload Me


    
End Sub


