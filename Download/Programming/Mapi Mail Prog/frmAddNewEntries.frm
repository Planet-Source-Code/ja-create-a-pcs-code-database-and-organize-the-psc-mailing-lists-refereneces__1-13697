VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddNewEntries 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add new Entries"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Prompt and add"
      Enabled         =   0   'False
      Height          =   380
      Left            =   1530
      TabIndex        =   2
      Top             =   2565
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   380
      Left            =   3075
      TabIndex        =   1
      Top             =   2565
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Enabled         =   0   'False
      Height          =   380
      Left            =   4095
      TabIndex        =   0
      Top             =   2565
      Width           =   930
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   75
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   -45
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wizard progress"
      Height          =   2300
      Left            =   300
      TabIndex        =   3
      Top             =   90
      Width           =   4755
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   1425
         TabIndex        =   4
         Top             =   195
         Width           =   3225
      End
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   345
         Picture         =   "frmAddNewEntries.frx":0000
         Top             =   540
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select date"
      Height          =   2300
      Left            =   300
      TabIndex        =   5
      Top             =   90
      Width           =   4755
      Begin VB.CommandButton Command4 
         Caption         =   "&Start"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   1755
         Width           =   960
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1515
         TabIndex        =   6
         Top             =   660
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   609
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36788
      End
      Begin VB.Label Label3 
         Caption         =   "Click here to start the search"
         Height          =   225
         Left            =   930
         TabIndex        =   10
         Top             =   1830
         Width           =   2100
      End
      Begin VB.Label Label2 
         Caption         =   $"frmAddNewEntries.frx":1482
         Height          =   645
         Left            =   210
         TabIndex        =   8
         Top             =   1035
         Width           =   4200
      End
      Begin VB.Label Label1 
         Caption         =   "The last date that code references have been saved to the databse was :"
         Height          =   480
         Left            =   225
         TabIndex        =   7
         Top             =   240
         Width           =   4230
      End
   End
End
Attribute VB_Name = "frmAddNewEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EmptyDB As Boolean
Private EmptyInbox As Boolean
Private NewerMessages As Boolean
Private LastCodeDate As Date
Private LastMsgID As Long


Private Sub Command1_Click()
Unload Me

End Sub


Private Sub Command2_Click()
UniVarPrompt = False
UniVarDate = DTPicker1.Value
Form1.Show 1, Me

End Sub

Private Sub Command4_Click()
Frame2.Visible = False
Frame1.Visible = True
Me.Refresh
DisplayImage Image1, 102
List1.AddItem "Looking for messages in 'Inbox'..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
MAPISession1.SignOn
List1.AddItem "Signing in..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
MAPIMessages1.SessionID = MAPISession1.SessionID
DisplayImage Image1, 103
List1.AddItem "Fetching messages..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
MAPIMessages1.Fetch
List1.AddItem "Fetched " & MAPIMessages1.MsgCount & " messages"
List1.Refresh
List1.ListIndex = List1.ListCount - 1
DisplayImage Image1, 104
List1.AddItem "Searching for messages..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
If MAPIMessages1.MsgCount = 0 Then
EmptyInbox = True
Else
EmptyInbox = False
List1.AddItem "Searching for newer messages from PSC..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
For i% = MAPIMessages1.MsgCount - 1 To 0 Step -1
MAPIMessages1.MsgIndex = i%
    If MAPIMessages1.MsgOrigAddress = "MailingList@planet-source-code.com" Or MAPIMessages1.MsgOrigAddress = "exhfirewall@exhedra.com" Or MAPIMessages1.MsgOrigAddress = "EXHFIREWALL@planet-source-code.com" Or MAPIMessages1.MsgOrigAddress = "iippoli1@tampabay.rr.com" Or MAPIMessages1.MsgOrigAddress = "MailingList@exhedra.com" Then ' Or MAPIMessages1.MsgOrigAddress = "iani@tampabay.rr.com" Then
        If MAPIMessages1.MsgDateReceived > LastCodeDate Or MAPIMessages1.MsgID > LastMsgID Then
        NewerMessages = True
        Exit For
        Else
        NewerMessages = False
        Exit For
        End If
    End If
Next i%
End If
MAPISession1.SignOff
If EmptyInbox = True Then
List1.AddItem "The inbox is empty... No new messages to add code from them."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
DisplayImage Image1, 106
Command1.Enabled = True
Command1.SetFocus
Exit Sub
End If
If NewerMessages = False Then
List1.AddItem "The inbox has no newer messages..."
List1.Refresh
DisplayImage Image1, 106
List1.ListIndex = List1.ListCount - 1
Command1.Enabled = True
Command1.SetFocus
Exit Sub
End If

If NewerMessages = True Then
List1.AddItem "The inbox has  newer messages..."
List1.Refresh
DisplayImage Image1, 105
List1.ListIndex = List1.ListCount - 1
List1.AddItem "Click Add to add new code references"
List1.Refresh
List1.ListIndex = List1.ListCount - 1
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command2.SetFocus
End If

End Sub

Private Sub Form_Load()
Frame1.Visible = False
Frame2.Visible = True

DisplayImage Image1, 101
List1.AddItem "Searching the Database... Please wait"
List1.Refresh
List1.ListIndex = List1.ListCount - 1
Me.Show
Me.Refresh
Dim tDB As dao.Database
Dim tSet As dao.Recordset
Set tDB = Workspaces(0).OpenDatabase(App.Path & "\psc.mdb", False, True)
Set tSet = tDB.OpenRecordset("SaveData", dbOpenSnapshot)
If tSet.EOF And tSet.BOF Then
EmptyDB = True
LastMsgID = -1
LastCodeDate = CDate("01/01/1900")
Else
EmptyDB = False
LastCodeDate = tSet.Fields("LastLisdateAccessed").Value
LastMsgID = tSet.Fields("LastMsgID").Value
End If
List1.AddItem "Finished searching the database"
List1.ListIndex = List1.ListCount - 1
Me.DTPicker1.Value = LastCodeDate

        

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmCodeEntries.Data1.Refresh


End Sub

