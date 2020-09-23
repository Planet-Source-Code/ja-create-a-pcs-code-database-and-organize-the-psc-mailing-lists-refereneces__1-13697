VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "GetMessages"
   ClientHeight    =   6210
   ClientLeft      =   1950
   ClientTop       =   1140
   ClientWidth     =   8430
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8430
   Begin MSComctlLib.StatusBar st 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   5865
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   2040
      Top             =   5625
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
      Left            =   7845
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Results"
      Height          =   5430
      Left            =   45
      TabIndex        =   1
      Top             =   15
      Width           =   8220
      Begin VB.Frame Frame5 
         Height          =   2055
         Left            =   2385
         TabIndex        =   32
         Top             =   750
         Visible         =   0   'False
         Width           =   3465
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "This mailing list may be in an attached HTML file. please wait until it is loaded..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            Left            =   390
            TabIndex        =   33
            Top             =   255
            Width           =   2745
         End
      End
      Begin SHDocVwCtl.WebBrowser Web1 
         Height          =   2775
         Left            =   1830
         TabIndex        =   31
         Top             =   7000
         Width           =   1635
         ExtentX         =   2884
         ExtentY         =   4895
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.TextBox Text1 
         Height          =   3465
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   210
         Width           =   7905
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   240
         Left            =   2310
         TabIndex        =   15
         Top             =   4485
         Visible         =   0   'False
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Current Listdate: "
         Height          =   195
         Left            =   105
         TabIndex        =   30
         Top             =   4200
         Width           =   2790
      End
      Begin VB.Label Label16 
         Caption         =   "Time passed: "
         Height          =   240
         Left            =   105
         TabIndex        =   29
         Top             =   5085
         Width           =   4485
      End
      Begin VB.Label Label15 
         Caption         =   "Time Left (estimated):"
         Height          =   240
         Left            =   105
         TabIndex        =   28
         Top             =   4755
         Width           =   4485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Progress:"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   4485
         Width           =   4485
      End
      Begin VB.Label Label7 
         Caption         =   "No code references added yet"
         Height          =   240
         Left            =   105
         TabIndex        =   16
         Top             =   3765
         Width           =   4485
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Results"
      Height          =   5430
      Left            =   30
      TabIndex        =   3
      Top             =   45
      Width           =   8220
      Begin VB.CheckBox Check1 
         Caption         =   "Prompt before adding a new code refrerence"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   3570
         Width           =   4155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   390
         Left            =   3975
         TabIndex        =   6
         Top             =   4125
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Height          =   390
         Left            =   5520
         TabIndex        =   5
         Top             =   4110
         Width           =   1395
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   2865
         TabIndex        =   4
         Top             =   900
         Width           =   4185
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         ForeColor       =   &H00FFFFFF&
         Height          =   2595
         Left            =   2865
         TabIndex        =   14
         Top             =   900
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.Image Image1 
         Height          =   3810
         Left            =   285
         Picture         =   "Form1.frx":0442
         Stretch         =   -1  'True
         Top             =   945
         Width           =   2085
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Results"
      Height          =   5430
      Left            =   30
      TabIndex        =   18
      Top             =   45
      Visible         =   0   'False
      Width           =   8220
      Begin VB.CommandButton Command1 
         Caption         =   "&Finish"
         Height          =   480
         Left            =   5340
         TabIndex        =   24
         Top             =   4155
         Width           =   1665
      End
      Begin VB.Label Label14 
         Caption         =   "Last Code reference ListDate='10/10/2000'"
         Height          =   240
         Left            =   2835
         TabIndex        =   27
         Top             =   3385
         Width           =   4785
      End
      Begin VB.Label Label13 
         Caption         =   "Last Code reference MSGID=123"
         Height          =   240
         Left            =   2835
         TabIndex        =   26
         Top             =   3023
         Width           =   4785
      End
      Begin VB.Label Label12 
         Caption         =   "Hit finish to close the wizard"
         Height          =   240
         Left            =   2835
         TabIndex        =   25
         Top             =   3750
         Width           =   2250
      End
      Begin VB.Label Label11 
         Caption         =   "In 3433 seconds"
         Height          =   240
         Left            =   2835
         TabIndex        =   23
         Top             =   2661
         Width           =   4800
      End
      Begin VB.Label Label10 
         Caption         =   "Added 43453 code references to database"
         Height          =   240
         Left            =   2835
         TabIndex        =   22
         Top             =   2299
         Width           =   4800
      End
      Begin VB.Label Label9 
         Caption         =   "Total 45454 mesages in 'Inbox'"
         Height          =   240
         Left            =   2835
         TabIndex        =   21
         Top             =   1575
         Width           =   4800
      End
      Begin VB.Label Label8 
         Caption         =   "Found and analyzed 245 messages from PCS"
         Height          =   240
         Left            =   2835
         TabIndex        =   20
         Top             =   1937
         Width           =   4800
      End
      Begin VB.Label Label6 
         Caption         =   "Finished adding new code references"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2685
         TabIndex        =   19
         Top             =   960
         Width           =   4905
      End
      Begin VB.Image Image3 
         Height          =   3810
         Left            =   285
         Picture         =   "Form1.frx":18C4
         Stretch         =   -1  'True
         Top             =   945
         Width           =   2085
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select date"
      Height          =   5430
      Left            =   75
      TabIndex        =   8
      Top             =   45
      Width           =   8220
      Begin VB.CommandButton Command4 
         Caption         =   "&Start"
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   3240
         Width           =   1590
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   4395
         TabIndex        =   10
         Top             =   1560
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   609
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36788
      End
      Begin VB.Image Image2 
         Height          =   3810
         Left            =   285
         Picture         =   "Form1.frx":2D46
         Stretch         =   -1  'True
         Top             =   945
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "The last date that code references have been saved to the databse was :"
         Height          =   480
         Left            =   3105
         TabIndex        =   13
         Top             =   1140
         Width           =   4230
      End
      Begin VB.Label Label2 
         Caption         =   $"Form1.frx":41C8
         Height          =   645
         Left            =   3090
         TabIndex        =   12
         Top             =   1935
         Width           =   4200
      End
      Begin VB.Label Label3 
         Caption         =   "Click here to start the search"
         Height          =   225
         Left            =   3420
         TabIndex        =   11
         Top             =   3345
         Width           =   2100
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CodeAddedCount As Long
Private Addresses() As String
Private CurentMessage As Integer

Private varWebBusy As Boolean

Function Exists_This_ListDate(lDate As Date) As Boolean
Dim stDB As DAO.Database
Dim stSet As DAO.Recordset
Dim t1 As String
Dim criteria As String
On Error GoTo ErrHandler

Set stDB = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb", False, True)
Set stSet = stDB.OpenRecordset("PSC1", dbOpenSnapshot)
criteria = "listdate=#" & Format(lDate, "mm/dd/yyyy") & "#"
stSet.FindFirst criteria
If Not stSet.NoMatch = True Then
Exists_This_ListDate = True
Else
Exists_This_ListDate = False
End If
stSet.Close
stDB.Close
Exit Function
ErrHandler:
Exists_This_ListDate = True
Exit Function
Resume Next
End Function

Function Found(H$, Title As String, dbd As DAO.Recordset) As Boolean
If dbd.EOF And dbd.BOF Then
Found = False
Exit Function
End If
If H$ = "" Then Found = False: Exit Function

criteria = "HTTP='" & H$ & "'"

dbd.FindFirst criteria

If dbd.NoMatch Then
Found = False
Else
MsgBox dbd.Fields("title").Value & vbCrLf & Title
Found = True
End If

End Function

Function GetETA(StartTime As Date, msgCount As Long, msgNow As Long) As Long
Dim TimeTillNow As Long
Dim TotalTime As Long

If msgNow = 0 Then
GetETA = -1
Exit Function
End If

TimeTillNow = Abs((DateDiff("s", StartTime, Now)))
TotalTime = msgCount * TimeTillNow / msgNow
GetETA = TotalTime - TimeTillNow

'GetETA = (msgCount - msgNow) * TimeTillNow / msgNow

End Function

Function GetListDate(note As String, EndLinePos As Integer) As Date

Dim DateLine As String
DateLine = GetLine(Text1.hwnd, 0)
Do
Y% = InStr(1, DateLine, ",")
EndLinePos = Len(DateLine)
If Y% > 0 Then
DateLine = Trim(Right$(DateLine, Len(DateLine) - Y%))
Else
Exit Do
End If
Loop



'For i% = 1 To Len(DateLine)
'X$ = Mid$(DateLine, i%, 10)
'If IsDate(X$) Then
If IsDate(DateLine) Then
GetListDate = MakeFormat(DateLine)
Else
GetListDate = VerifyListDate(Me, Left$(MAPIMessages1.MsgNoteText, 200), MAPIMessages1.msgID, MAPIMessages1.MsgDateReceived)
End If

'Exit Function
'End If
'Next i%
EndLinePos = 0

End Function

Function GetListFromAttachment() As String
varWebBusy = True
Dim cFile As String
If MAPIMessages1.AttachmentCount > 0 Then
MAPIMessages1.AttachmentIndex = 0
cFile = MAPIMessages1.AttachmentPathName
Web1.Navigate cFile
End If

FileLookUp:
Frame5.Visible = True
Refresh
S$ = st.SimpleText
st.SimpleText = S$ & " Searching in attached list"
Dim d As IHTMLDocument2
Set d = Web1.Document
If d.url <> CreateWebLocation(cFile) Then
Web1.Navigate cFile
DoEvents
st.SimpleText = S$ & " Waiting for attcheched file to load"
GoTo FileLookUp
End If
Frame5.Visible = False
Refresh
st.SimpleText = S$
GetListFromAttachment = d.Body.innerText
Set d = Nothing
Web1.Navigate App.path & "\Blank.htm"

End Function

Private Sub LoadAddresses()
Dim tSet As DAO.Recordset
Dim tdb As DAO.Database
Set tdb = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set tSet = tdb.OpenRecordset("Addresses", dbOpenSnapshot)
If tSet.EOF And tSet.BOF Then
ReDim Preserve Addresses(0)
Addresses(0) = ""
Else
ccount% = 0
    Do While Not tSet.EOF
    ReDim Preserve Addresses(ccount%)
    Addresses(ccount%) = tSet.Fields("Address").Value & vbNullString
    ccount% = ccount% + 1
    tSet.MoveNext
    Loop
End If
tSet.Close
Set tSet = Nothing
tdb.Close
Set tdb = Nothing


End Sub

Function MakeFormat(f As String) As Date
Dim cDay As Integer, cMonth As Integer, cYear As Long
cDay = CInt(Mid$(f, 4, 2))
cMonth = CInt(Left$(f, 2))
cYear = CLng(Right$(f, 4))
If cYear = 200 Then cYear = 2000
MakeFormat = CDate(cDay & "/" & cMonth & "/" & cYear)


End Function


Function msgOriginalAddress(ad As String) As Boolean
For i% = 0 To UBound(Addresses)
If ad = Addresses(i%) Then
    msgOriginalAddress = True
    Exit Function
End If
Next i%
msgOriginalAddress = False

End Function

Sub ReadCodeEntry(txt As TextBox, index As Integer)
Dim BeginLine As Integer
Dim EndLine As Integer
Dim HttpStart As Integer
Dim Des As String
CE.Accessed = 0
BeginLine = -1
CE.Category = "<UNKNOWN>"
CE.Compatibility = "<UNKNOWN>"
CE.DateSubmitted = 0
CE.http = ""
CE.ID = 0
CE.Level = "<UNKNOWN>"
CE.Text = ""
CE.Title = ""

For i% = 0 To GettxtLineCount(txt.hwnd) - 1
    L% = Len(CStr(index) & ")")
    m$ = GetLine(txt.hwnd, i%)
    If Left(m$, L%) = CStr(index) & ")" Then
    BeginLine = i%
    ElseIf (Left$(m$, 8) = "********" Or Left$(m$, 8) = "========") And BeginLine >= 0 Then
    EndLine = i%
    Exit For
    End If
Next i%
m$ = GetLine(txt.hwnd, BeginLine)
If Len(m$) > 0 Then
CE.Title = Right(m$, Len(m$) - L%)
Else
CE.Title = "No title"
End If

For i% = BeginLine + 1 To EndLine
    m$ = GetLine(txt.hwnd, i%)
    If UCase(Left$(m$, 9)) = "CATEGORY:" Then
    CE.Category = Trim(Right$(m$, Len(m$) - 9))
    End If
    
    If UCase(Left$(m$, 6)) = "LEVEL:" Then
    CE.Level = Trim(Right$(m$, Len(m$) - 6))
    End If
    
    If Left$(m$, 33) = "http://www.planet-source-code.com" Then
    CE.http = m$
    HttpStart = i% - 1
    End If
    
    If UCase(Left$(m$, 14)) = "COMPATIBILITY:" Then
    CE.Compatibility = Trim(Right$(m$, Len(m$) - 14))
    End If
    
    If UCase(Left$(m$, 12)) = "SUBMITTED ON" Then
    If IsDate(Trim(Mid$(m$, 13, 10))) Then
    CE.DateSubmitted = CDate(Trim(Mid$(m$, 13, 10)))
    Else
        If IsDate(Trim(Mid$(m$, 13, 8))) Then
        CE.DateSubmitted = CDate(Trim(Mid$(m$, 13, 8)))
        Else
        CE.DateSubmitted = 0
        End If
    End If
    
    f$ = Left$(m$, Len(m$) - 7)
    S% = InStr(1, f$, "accessed") + Len("accessed")
    CE.Accessed = Val(Right$(f$, Len(f$) - S%))
    
    
    End If
    
    
    
    
Next i%
    For i% = BeginLine To EndLine
    m$ = GetLine(txt.hwnd, i%)
    If UCase(Left$(m$, 12)) = "DESCRIPTION:" Then
    For t% = i% To HttpStart
    Des = Des & vbCrLf & GetLine(txt.hwnd, t%)
    Next t%
    If Len(Des) > 0 Then
    CE.Text = Right$(Des, Len(Des) - 12)
    If Right$(CE.Text, 27) = "Complete source code is at:" Then
    CE.Text = Left$(CE.Text, Len(CE.Text) - 27)
    End If
    If Left$(CE.Text, 3) = "n: " Then
    CE.Text = Right$(CE.Text, Len(CE.Text) - 3)
    End If
    
    Else
    CE.Text = "No description"
    End If
    End If
    Next i%
    

End Sub

Sub SaveCodeEntry()
Dim dbs As DAO.Database
Dim dbset As DAO.Recordset
Dim cloneDBset  As DAO.Recordset
Set dbs = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set dbset = dbs.OpenRecordset("PSC1", dbOpenDynaset)
Set cloneDBset = dbset.Clone

    If UniVarPrompt = True Then
        Dim d As PromptRespone
        d = GetPromptReponce
        If d = yes Then
            dbset.AddNew
            dbset.Fields("msgID").Value = CE.ID
            dbset.Fields("Title").Value = CE.Title
            dbset.Fields("Category").Value = CE.Category
            dbset.Fields("Level").Value = CE.Level
            dbset.Fields("Description").Value = CE.Text
            dbset.Fields("compatibility").Value = CE.Compatibility
            dbset.Fields("DateSumbitted").Value = CE.DateSubmitted
            dbset.Fields("Listdate").Value = CE.ListDate
            dbset.Fields("Http").Value = CE.http
            dbset.Fields("Accessed").Value = CE.Accessed
            dbset.Fields("[CodeOfTheDay]").Value = CE.CodeOfTheDay
            dbset.Update
            LastDateToStore = CE.ListDate
            LastMEssageIDtoStore = CE.ID
            CodeAddedCount = CodeAddedCount + 1
        ElseIf d = No Then
        Else
            UniVarPrompt = False
            dbset.AddNew
            dbset.Fields("msgID").Value = CE.ID
            dbset.Fields("Title").Value = CE.Title
            dbset.Fields("Category").Value = CE.Category
            dbset.Fields("Level").Value = CE.Level
            dbset.Fields("Description").Value = CE.Text
            dbset.Fields("compatibility").Value = CE.Compatibility
            dbset.Fields("DateSumbitted").Value = CE.DateSubmitted
            dbset.Fields("Listdate").Value = CE.ListDate
            dbset.Fields("Http").Value = CE.http
            dbset.Fields("Accessed").Value = CE.Accessed
            dbset.Fields("[CodeOfTheDay]").Value = CE.CodeOfTheDay
            dbset.Update
            LastDateToStore = CE.ListDate
            LastMEssageIDtoStore = CE.ID
            CodeAddedCount = CodeAddedCount + 1
        End If
    Else
        dbset.AddNew
        dbset.Fields("msgID").Value = CE.ID
        dbset.Fields("Title").Value = CE.Title
        dbset.Fields("Category").Value = CE.Category
        dbset.Fields("Level").Value = CE.Level
        dbset.Fields("Description").Value = CE.Text
        dbset.Fields("compatibility").Value = CE.Compatibility
        dbset.Fields("DateSumbitted").Value = CE.DateSubmitted
        dbset.Fields("Listdate").Value = CE.ListDate
        dbset.Fields("Http").Value = CE.http
        dbset.Fields("Accessed").Value = CE.Accessed
        dbset.Fields("[CodeOfTheDay]").Value = CE.CodeOfTheDay
        dbset.Update
        LastDateToStore = CE.ListDate
        LastMEssageIDtoStore = CE.ID
        CodeAddedCount = CodeAddedCount + 1
    End If


Refresh

dbset.Close
dbs.Close
Label7.Caption = "Added " & CodeAddedCount & " code references"
End Sub


Sub SaveLastData(d As Date, ID As Long)
Dim dbs As DAO.Database
Dim dbset As DAO.Recordset
Set dbs = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set dbset = dbs.OpenRecordset("SaveData", dbOpenDynaset)
If dbset.EOF And dbset.BOF Then
dbset.AddNew
Else
dbset.Edit
End If
dbset.Fields("LastLisdateAccessed").Value = d
dbset.Fields("LastMsgID").Value = ID
dbset.Update
dbset.Close
dbs.Close

End Sub


Private Sub Command1_Click()
Unload Me
frmCodeEntries.Data1.Refresh
frmCodeEntries.Data1.Recordset.MoveLast

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim pr As Single
Dim StartTime As Date, FinishTime As Date, TotalTime As Long
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
UniVarDate = DTPicker1.Value
UniVarPrompt = CBool(Check1.Value)
pb1.Min = 1
pb1.Max = MAPIMessages1.msgCount
pb1.Value = 1
pb1.Visible = True
Dim ListDate As Date
Dim CodeEntriesCount As Integer
Dim msgLineCount As Long
Dim EndLinePos As Integer
msgCount% = 0
StartTime = Now
For i% = 0 To MAPIMessages1.msgCount - 1
Refresh
Label15.Caption = "Estimated time left: " & FormatTime(GetETA(StartTime, MAPIMessages1.msgCount, CLng(i%)))
Label16.Caption = "Elapsed time: " & FormatTime(DateDiff("s", StartTime, Now))
pb1.Value = i% + 1
pr = 100 * CLng(i%)
pr = pr / MAPIMessages1.msgCount
Label5.Caption = "Progress : " & CStr(CInt(pr)) & "%"

    MAPIMessages1.MsgIndex = i%
    st.SimpleText = "Analyzing message: " & MAPIMessages1.MsgIndex & " of " & MAPIMessages1.msgCount
    If msgOriginalAddress(MAPIMessages1.MsgOrigAddress) = True Then
        msgCount% = msgCount% + 1
        
        st.SimpleText = st.SimpleText & ", found " & msgCount% & " messages from PSC."
        If InStr(1, UCase(MAPIMessages1.MsgSubject), "CODE OF THE DAY") <> 0 Then
        If Left$(MAPIMessages1.MsgNoteText, Len(MAPIMessages1.MsgSubject)) = MAPIMessages1.MsgSubject Then
        Text1.Text = Right$(MAPIMessages1.MsgNoteText, Len(MAPIMessages1.MsgNoteText) - Len(MAPIMessages1.MsgSubject))
        Else
        Text1.Text = MAPIMessages1.MsgNoteText
        End If
        If Trim(Text1.Text) = "" Then
        'check to see if there are any attachments
        Text1.Text = GetListFromAttachment
        st.SimpleText = st.SimpleText & " Checking for attachments"
        End If
        
        'Text1.Refresh
        ClearEmptyLines Text1
        'Text1.Refresh
        EndLinePos = 0
        ListDate = GetListDate(Text1.Text, EndLinePos)
        Label17.Caption = "Current Listdate: " & ListDate
            If ListDate > UniVarDate Then
                If Exists_This_ListDate(ListDate) = False Then
                    If Trim(Text1.Text) <> "" Then
                    Text1.Text = Right$(Text1.Text, Len(Text1.Text) - EndLinePos - 2)
                    Text1.Refresh
                    msgLineCount = GettxtLineCount(Text1.hwnd)
                    msgLineCount = RemoveLinesUntilString(Text1, "Code of the Day:", "Today's code is:")
                    Text1.Refresh
                Dim LineNo As Integer
                    CodeEntriesCount = GetCodeEntriesCount(Text1.hwnd, Text1.Text, LineNo)
                    msgLineCount = RemoveLinesUntilString(Text1, "******", "=======")
                    Text1.Refresh
                    For L = 0 To 10
                    If Left(GetLine(Text1.hwnd, L), 1) = "1" Then
                        If L > 0 Then
                        m$ = GetLine(Text1.hwnd, L - 1)
                        msgLineCount = RemoveLinesUntilString(Text1, m$, m$)
                        Exit For
                        End If
                    End If
                    Next L
                    For t% = 1 To CodeEntriesCount
                        ReadCodeEntry Text1, t%
                        CE.ListDate = ListDate
                        If t% = 1 Then
                        CE.CodeOfTheDay = True
                        Else
                        CE.CodeOfTheDay = False
                        End If
                        CE.ID = MAPIMessages1.msgID
                        SaveCodeEntry
                    Next t%
                    End If
                End If
            End If
        End If
    End If
Next i%
If CodeAddedCount > 0 Then SaveLastData LastDateToStore, LastMEssageIDtoStore


pb1.Visible = False
st.SimpleText = "Fininshed analyzing " & MAPIMessages1.msgCount & " messages of " & MAPIMessages1.msgCount & ". Found " & msgCount% & " messages." & " Current Listdate: " & CE.ListDate
FinishTime = Now
TotalTime = Abs(DateDiff("s", FinishTime, StartTime))
Label9.Caption = "Total " & MAPIMessages1.msgCount & " messages in Inbox."
Label8.Caption = "Found and analyzed " & msgCount & " messages from PSC."
Label10.Caption = "Added " & CodeAddedCount & " code references,"
Label11.Caption = "in " & FormatTime(TotalTime) & " seconds."
Label13.Caption = "Last code reference MsgID = " & LastMEssageIDtoStore
Label14.Caption = "Last code reference listdate = " & LastDateToStore



Frame4.Visible = True
Frame1.Visible = False

MousePointer = 0

End Sub
Private Sub Command4_Click()
On Error GoTo ErrH
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
MousePointer = 11
Me.Refresh
DisplayImage Image1, 102
List1.AddItem "Looking for messages in 'Inbox'..."
st.SimpleText = "Looking for messages in 'Inbox'..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
MAPISession1.SignOn
List1.AddItem "Signing in..."
st.SimpleText = "Signing in..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
MAPIMessages1.SessionID = MAPISession1.SessionID
DisplayImage Image1, 103
List1.AddItem "Fetching messages..."
st.SimpleText = "Fetching messages..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
MAPIMessages1.Fetch
List1.AddItem "Fetched " & MAPIMessages1.msgCount & " messages"
st.SimpleText = "Fetched " & MAPIMessages1.msgCount & " messages"
List1.Refresh
List1.ListIndex = List1.ListCount - 1
DisplayImage Image1, 104
List1.AddItem "Searching for messages..."
st.SimpleText = "Searching for messages..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
If MAPIMessages1.msgCount = 0 Then
EmptyInbox = True
Else
EmptyInbox = False
List1.AddItem "Searching for newer messages from PSC..."
st.SimpleText = "Searching for newer messages from PSC..."
List1.Refresh
List1.ListIndex = List1.ListCount - 1
For i% = 0 To MAPIMessages1.msgCount - 1
MAPIMessages1.MsgIndex = i%
    If MAPIMessages1.MsgOrigAddress = "MailingList@planet-source-code.com" Or MAPIMessages1.MsgOrigAddress = "exhfirewall@exhedra.com" Or MAPIMessages1.MsgOrigAddress = "EXHFIREWALL@planet-source-code.com" Or MAPIMessages1.MsgOrigAddress = "iippoli1@tampabay.rr.com" Or MAPIMessages1.MsgOrigAddress = "MailingList@exhedra.com" Then ' Or MAPIMessages1.MsgOrigAddress = "iani@tampabay.rr.com" Then
        If MAPIMessages1.MsgDateReceived > LastCodeDate Or MAPIMessages1.msgID > LastMsgID Then
        NewerMessages = True
        Exit For
        Else
        NewerMessages = False
        Exit For
        End If
    End If
Next i%
End If
If EmptyInbox = True Then
List1.AddItem "The inbox is empty... No new messages to add code from them."
st.SimpleText = "Ready"
List1.Refresh
List1.ListIndex = List1.ListCount - 1
DisplayImage Image1, 106
Command3.Enabled = False
Command2.Enabled = True
Check1.Enabled = False
Command2.SetFocus
MousePointer = 0
List1.Visible = False
Label4.Caption = "The inbox is empty... No new messages to add code from them."
Label4.Visible = True

Exit Sub
End If
If NewerMessages = False Then
List1.AddItem "The inbox has no newer messages..."
st.SimpleText = "Ready"
List1.Refresh
DisplayImage Image1, 106
List1.ListIndex = List1.ListCount - 1
Command3.Enabled = False
Command2.Enabled = True
Check1.Enabled = False
Command2.SetFocus
MousePointer = 0
List1.Visible = False
Label4.Caption = "The inbox has no newer messages..."
Label4.Visible = True

Exit Sub
End If

If NewerMessages = True Then
List1.AddItem "The inbox has  newer messages..."
st.SimpleText = "Ready"
List1.Refresh
DisplayImage Image1, 105
List1.ListIndex = List1.ListCount - 1
List1.AddItem "Click Add to add new code references"
List1.Refresh
List1.ListIndex = List1.ListCount - 1
Command3.Enabled = True
Command2.Enabled = True
Check1.Enabled = True
Command3.SetFocus
List1.Visible = False
Label4.Caption = "The 'Inbox' contains new messages from PSC" & vbCrLf & "Click 'Add' to add new code references"
Label4.Visible = True
MousePointer = 0

End If
Exit Sub
ErrH:
MsgBox Err.Source
Resume Next
End Sub



Private Sub Form_Load()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Refresh
DisplayImage Image2, 101
Me.st.SimpleText = "Load valid eMail addresses..."
Me.Refresh
LoadAddresses
Me.st.SimpleText = "Searching the Database... Please wait"
Me.Refresh
Dim tdb As DAO.Database
Dim tSet As DAO.Recordset
Set tdb = Workspaces(0).OpenDatabase(App.path & "\psc.mdb", False, True)
Set tSet = tdb.OpenRecordset("SaveData", dbOpenSnapshot)
If tSet.EOF And tSet.BOF Then
EmptyDB = True
LastMsgID = -1
LastCodeDate = CDate("01/01/1900")
Else
EmptyDB = False
LastCodeDate = tSet.Fields("LastLisdateAccessed").Value
LastMsgID = tSet.Fields("LastMsgID").Value
End If
st.SimpleText = "Finished searching the database"
Me.DTPicker1.Value = LastCodeDate
Refresh




End Sub


Private Sub Form_Unload(Cancel As Integer)
If MAPISession1.SessionID > 0 Then
MAPISession1.SignOff
End If

End Sub






Private Sub Web1_DocumentComplete(ByVal pDisp As Object, url As Variant)
varWebBusy = False
End Sub

