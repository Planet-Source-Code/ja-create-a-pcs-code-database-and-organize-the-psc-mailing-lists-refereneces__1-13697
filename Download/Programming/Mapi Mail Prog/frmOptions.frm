VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   7350
   ClientLeft      =   555
   ClientTop       =   735
   ClientWidth     =   8100
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Valid PSC eMail original addresses"
      Height          =   1920
      Left            =   90
      TabIndex        =   56
      Top             =   4545
      Width           =   7815
      Begin VB.CommandButton Command8 
         Caption         =   "Find eMails"
         Height          =   300
         Left            =   180
         TabIndex        =   61
         Top             =   1395
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Remove A&ll"
         Height          =   300
         Left            =   180
         TabIndex        =   60
         Top             =   1050
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Remove"
         Height          =   300
         Left            =   165
         TabIndex        =   59
         Top             =   690
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Add"
         Height          =   300
         Left            =   165
         TabIndex        =   58
         Top             =   345
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   1605
         TabIndex        =   57
         Top             =   300
         Width           =   6030
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7365
      Top             =   6675
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   6855
      TabIndex        =   51
      Top             =   6765
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   5400
      TabIndex        =   50
      Top             =   6765
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "Startup Filter"
      Height          =   1350
      Left            =   90
      TabIndex        =   49
      Top             =   150
      Width           =   7815
      Begin VB.CommandButton Command4 
         Caption         =   "Open filter"
         Enabled         =   0   'False
         Height          =   375
         Left            =   210
         TabIndex        =   55
         Top             =   750
         Width           =   1020
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   630
         Left            =   1275
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Text            =   "frmOptions.frx":0442
         Top             =   600
         Width           =   6315
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Load a specified filter on startup"
         Height          =   225
         Left            =   1275
         TabIndex        =   53
         Top             =   300
         Width           =   6255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customize the appearence of the search box"
      Height          =   2865
      Left            =   90
      TabIndex        =   0
      Top             =   1545
      Width           =   7815
      Begin VB.CommandButton Command3 
         Caption         =   "Set Defaults"
         Height          =   270
         Left            =   1665
         TabIndex        =   52
         Top             =   2355
         Width           =   1950
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2685
         ScaleHeight     =   330
         ScaleWidth      =   1005
         TabIndex        =   45
         Top             =   1935
         Width           =   1005
         Begin VB.OptionButton lftComment 
            Height          =   315
            Left            =   0
            Picture         =   "frmOptions.frx":0494
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton cntComment 
            Height          =   315
            Left            =   315
            Picture         =   "frmOptions.frx":0596
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton rtComment 
            Height          =   315
            Left            =   645
            Picture         =   "frmOptions.frx":0698
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2685
         ScaleHeight     =   330
         ScaleWidth      =   1005
         TabIndex        =   41
         Top             =   1515
         Width           =   1005
         Begin VB.OptionButton lftHTTP 
            Height          =   315
            Left            =   0
            Picture         =   "frmOptions.frx":079A
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton cntHTTP 
            Height          =   315
            Left            =   315
            Picture         =   "frmOptions.frx":089C
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton rtHTTP 
            Height          =   315
            Left            =   645
            Picture         =   "frmOptions.frx":099E
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2685
         ScaleHeight     =   330
         ScaleWidth      =   1005
         TabIndex        =   37
         Top             =   1110
         Width           =   1005
         Begin VB.OptionButton lftText 
            Height          =   315
            Left            =   0
            Picture         =   "frmOptions.frx":0AA0
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton cntText 
            Height          =   315
            Left            =   315
            Picture         =   "frmOptions.frx":0BA2
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton rtText 
            Height          =   315
            Left            =   645
            Picture         =   "frmOptions.frx":0CA4
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2685
         ScaleHeight     =   330
         ScaleWidth      =   1005
         TabIndex        =   33
         Top             =   690
         Width           =   1005
         Begin VB.OptionButton lftField 
            Height          =   315
            Left            =   0
            Picture         =   "frmOptions.frx":0DA6
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton cntField 
            Height          =   315
            Left            =   315
            Picture         =   "frmOptions.frx":0EA8
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton rtField 
            Height          =   315
            Left            =   645
            Picture         =   "frmOptions.frx":0FAA
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2685
         ScaleHeight     =   330
         ScaleWidth      =   1005
         TabIndex        =   29
         Top             =   285
         Width           =   1005
         Begin VB.OptionButton rtTitle 
            Height          =   315
            Left            =   645
            Picture         =   "frmOptions.frx":10AC
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton cntTitle 
            Height          =   315
            Left            =   315
            Picture         =   "frmOptions.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   0
            Width           =   300
         End
         Begin VB.OptionButton lftTitle 
            Height          =   315
            Left            =   0
            Picture         =   "frmOptions.frx":12B0
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.CheckBox bldComment 
         Height          =   315
         Left            =   1680
         Picture         =   "frmOptions.frx":13B2
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1935
         Width           =   300
      End
      Begin VB.CheckBox itComment 
         Height          =   315
         Left            =   2010
         Picture         =   "frmOptions.frx":14B4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1935
         Width           =   300
      End
      Begin VB.CheckBox uComment 
         Height          =   315
         Left            =   2340
         Picture         =   "frmOptions.frx":15B6
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1935
         Width           =   300
      End
      Begin VB.CheckBox bldHTTP 
         Height          =   315
         Left            =   1680
         Picture         =   "frmOptions.frx":16B8
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1515
         Width           =   300
      End
      Begin VB.CheckBox itHTTP 
         Height          =   315
         Left            =   2010
         Picture         =   "frmOptions.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1515
         Width           =   300
      End
      Begin VB.CheckBox uHTTP 
         Height          =   315
         Left            =   2340
         Picture         =   "frmOptions.frx":18BC
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1515
         Width           =   300
      End
      Begin VB.CheckBox bldText 
         Height          =   315
         Left            =   1680
         Picture         =   "frmOptions.frx":19BE
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1110
         Width           =   300
      End
      Begin VB.CheckBox itText 
         Height          =   315
         Left            =   2010
         Picture         =   "frmOptions.frx":1AC0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1110
         Width           =   300
      End
      Begin VB.CheckBox uText 
         Height          =   315
         Left            =   2340
         Picture         =   "frmOptions.frx":1BC2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1110
         Width           =   300
      End
      Begin VB.CheckBox bldField 
         Height          =   315
         Left            =   1680
         Picture         =   "frmOptions.frx":1CC4
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   690
         Width           =   300
      End
      Begin VB.CheckBox itField 
         Height          =   315
         Left            =   2010
         Picture         =   "frmOptions.frx":1DC6
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   690
         Width           =   300
      End
      Begin VB.CheckBox uField 
         Height          =   315
         Left            =   2340
         Picture         =   "frmOptions.frx":1EC8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   690
         Width           =   300
      End
      Begin VB.CheckBox uTitle 
         Height          =   315
         Left            =   2340
         Picture         =   "frmOptions.frx":1FCA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   285
         Width           =   300
      End
      Begin VB.CheckBox itTitle 
         Height          =   315
         Left            =   2010
         Picture         =   "frmOptions.frx":20CC
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   285
         Width           =   300
      End
      Begin VB.CheckBox bldTitle 
         Height          =   315
         Left            =   1680
         Picture         =   "frmOptions.frx":21CE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   285
         Width           =   300
      End
      Begin RichTextLib.RichTextBox rt1 
         Height          =   2670
         Left            =   4050
         TabIndex        =   12
         Top             =   150
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4710
         _Version        =   393217
         Enabled         =   0   'False
         FileName        =   "D:\Download\Programming\Mapi Mail Prog\Sample.rtf"
         TextRTF         =   $"frmOptions.frx":22D0
      End
      Begin Project1.ArielColorBox cbTitle 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   285
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Palette         =   5
      End
      Begin Project1.ArielColorBox cbField 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   690
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Palette         =   5
      End
      Begin Project1.ArielColorBox cbText 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1110
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Palette         =   5
      End
      Begin Project1.ArielColorBox cbHTTP 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1515
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Palette         =   5
      End
      Begin Project1.ArielColorBox cbComment 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   1935
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Palette         =   5
      End
      Begin Project1.ArielColorBox cbBack 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   2340
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Palette         =   5
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Background"
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   11
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   1986
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "HTTP"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   1572
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Text"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   1158
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Field"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   744
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Title"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   330
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents ArTimer As ArielTimer  'Must be private to raise events
Attribute ArTimer.VB_VarHelpID = -1

Private Const CommentStart = 0
Private Const Commentlength = 7
Private Const TitleStart = 9
Private Const TitleLength = 5

Sub LoadAddressList()
List1.Clear
Dim tSet As DAO.Recordset
Dim tdb As DAO.Database
Set tdb = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set tSet = tdb.OpenRecordset("Addresses", dbOpenSnapshot)
If tSet.BOF And tSet.EOF Then
Else
    Do While Not tSet.EOF
    List1.AddItem tSet.Fields("Address").Value & vbNullString
    tSet.MoveNext
    Loop
End If
tSet.Close
Set tSet = Nothing
tdb.Close
Set tdb = Nothing

End Sub


Sub SaveAddressList()
Dim tSet As DAO.Recordset
Dim tdb As DAO.Database
Set tdb = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set tSet = tdb.OpenRecordset("Addresses", dbOpenDynaset)
Do While Not tSet.EOF
tSet.Delete
tSet.MoveNext
Loop
For i% = 0 To List1.ListCount - 1
tSet.AddNew
tSet.Fields("Address").Value = List1.List(i%)
tSet.Update
Next i%

tSet.Close
Set tSet = Nothing
tdb.Close
Set tdb = Nothing

End Sub

Private Sub bldComment_Click()
rt1.SelStart = CommentStart
rt1.SelLength = Commentlength
rt1.SelBold = CBool(bldComment.Value)
rt1.SelLength = 0
End Sub

Private Sub bldField_Click()
For i% = 0 To 99 Step 33
rt1.SelStart = 16 + i%
rt1.SelLength = 6
rt1.SelBold = CBool(bldField.Value)
rt1.SelLength = 0
Next i%
rt1.SelStart = 148
rt1.SelLength = 9
rt1.SelBold = CBool(bldField.Value)
rt1.SelLength = 0

rt1.SelStart = 164
rt1.SelLength = 10
rt1.SelBold = CBool(bldField.Value)
rt1.SelLength = 0

rt1.SelStart = 201
rt1.SelLength = 9
rt1.SelBold = CBool(bldField.Value)
rt1.SelLength = 0


End Sub

Private Sub bldHTTP_Click()
Dim L As Boolean
L = CBool(bldHTTP)
rt1.SelStart = 157
rt1.SelLength = 5
rt1.SelBold = L

rt1.SelStart = 174
rt1.SelLength = 25
rt1.SelBold = L

rt1.SelStart = 210
rt1.SelLength = 25
rt1.SelBold = L

rt1.SelLength = 0
End Sub

Private Sub bldText_Click()
rt1.SelStart = 22
rt1.SelLength = 26
rt1.SelBold = CBool(bldText.Value)
rt1.SelLength = 0

rt1.SelStart = 55
rt1.SelLength = 26
rt1.SelBold = CBool(bldText.Value)
rt1.SelLength = 0

rt1.SelStart = 88
rt1.SelLength = 26
rt1.SelBold = CBool(bldText.Value)
rt1.SelLength = 0

rt1.SelStart = 121
rt1.SelLength = 26
rt1.SelBold = CBool(bldText.Value)
rt1.SelLength = 0


End Sub


Private Sub bldTitle_Click()
rt1.SelStart = TitleStart
rt1.SelLength = TitleLength
rt1.SelBold = CBool(bldTitle.Value)
rt1.SelLength = 0

End Sub

Private Sub cbComment_Click()
rt1.SelStart = CommentStart
rt1.SelLength = Commentlength
rt1.SelColor = cbComment.SelectedColor
rt1.SelLength = 0

End Sub

Private Sub cbField_Click()
For i% = 0 To 99 Step 33
rt1.SelStart = 16 + i%
rt1.SelLength = 6
rt1.SelColor = cbField.SelectedColor
Next i%
rt1.SelStart = 148
rt1.SelLength = 9
rt1.SelColor = cbField.SelectedColor

rt1.SelStart = 164
rt1.SelLength = 10
rt1.SelColor = cbField.SelectedColor

rt1.SelStart = 201
rt1.SelLength = 9
rt1.SelColor = cbField.SelectedColor
rt1.SelLength = 0

End Sub

Private Sub cbHTTP_Click()
rt1.SelStart = 157
rt1.SelLength = 5
rt1.SelColor = cbHTTP.SelectedColor

rt1.SelStart = 174
rt1.SelLength = 25
rt1.SelColor = cbHTTP.SelectedColor

rt1.SelStart = 210
rt1.SelLength = 25
rt1.SelColor = cbHTTP.SelectedColor

rt1.SelLength = 0

End Sub

Private Sub cbText_Click()
rt1.SelStart = 22
rt1.SelLength = 26
rt1.SelColor = cbText.SelectedColor

rt1.SelStart = 55
rt1.SelLength = 26
rt1.SelColor = cbText.SelectedColor

rt1.SelStart = 88
rt1.SelLength = 26
rt1.SelColor = cbText.SelectedColor

rt1.SelStart = 121
rt1.SelLength = 26
rt1.SelColor = cbText.SelectedColor
rt1.SelLength = 0

End Sub

Private Sub cbTitle_Click()
rt1.SelStart = TitleStart
rt1.SelLength = TitleLength
rt1.SelColor = cbTitle.SelectedColor
rt1.SelLength = 0
End Sub


Private Sub Check1_Click()
Command4.Enabled = CBool(check1.Value)
Text1.Enabled = CBool(check1.Value)
End Sub

Private Sub cntComment_Click()
rt1.SelStart = CommentStart
rt1.SelLength = Commentlength
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0

End Sub

Private Sub cntField_Click()
For i% = 0 To 99 Step 33
rt1.SelStart = 16 + i%
rt1.SelLength = 6
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0
Next i%
rt1.SelStart = 148
rt1.SelLength = 9
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0

rt1.SelStart = 164
rt1.SelLength = 10
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0

rt1.SelStart = 201
rt1.SelLength = 9
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0

End Sub

Private Sub cntHTTP_Click()
rt1.SelStart = 157
rt1.SelLength = 5
rt1.SelAlignment = rtfCenter

rt1.SelStart = 174
rt1.SelLength = 25
rt1.SelAlignment = rtfCenter

rt1.SelStart = 210
rt1.SelLength = 25
rt1.SelAlignment = rtfCenter

rt1.SelLength = 0

End Sub

Private Sub cntText_Click()
rt1.SelStart = 22
rt1.SelLength = 26
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0

rt1.SelStart = 55
rt1.SelLength = 26
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0

rt1.SelStart = 88
rt1.SelLength = 26
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0

rt1.SelStart = 121
rt1.SelLength = 26
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0
End Sub

Private Sub cntTitle_Click()
rt1.SelStart = TitleStart
rt1.SelLength = TitleLength
rt1.SelAlignment = rtfCenter
rt1.SelLength = 0

End Sub

Private Sub Check11_Click()

End Sub

Private Sub cbBack_Click()
rt1.BackColor = cbBack.SelectedColor

End Sub

Private Sub Command1_Click()
frmMessageCheckKey = "ShowThisAgainOptions"
If CBool(GetSetting("PSC database", "Settings", frmMessageCheckKey, "-1")) = True Then
frmMessageButtonCaption = "&OK"
frmMessageCheckCaption = "Don't show this again"
frmMessageCheckVisible = True
frmMessageMessage = "The formatting toy have selected will be active only for the new searches, not for the new ones."
frmMessage.Show 1, Me
End If
'Misc
LoadSQLatStartup = CBool(check1.Value)
StartUpSQL = Text1.Text
PutRegString "LoadSQLAtStartUp", CStr(LoadSQLatStartup)
PutRegString "StartUpSQL", StartUpSQL

'Title
TitleColor = cbTitle.SelectedColor
TitleBold = CBool(bldTitle.Value)
TitleItalics = CBool(itTitle.Value)
TitleUnderline = CBool(uTitle.Value)
If lftTitle.Value = True Then
TitleAllingment = rtfLeft
ElseIf rtTitle.Value = True Then
TitleAllingment = rtfRight
Else
TitleAllingment = rtfCenter
End If
PutRegString "TitleColor", CStr(TitleColor)
PutRegString "TitleBold", CStr(TitleBold)
PutRegString "TitleItalics", CStr(TitleItalics)
PutRegString "TitleUnderline", CStr(TitleUnderline)
PutRegString "TitleAlign", CStr(TitleAllingment)

'Text
TextColor = cbText.SelectedColor
TextBold = CBool(bldText.Value)
TextItalics = CBool(itText.Value)
TextUnderline = CBool(uText.Value)
If lftText.Value = True Then
TextAllingment = rtfLeft
ElseIf rtText.Value = True Then
TextAllingment = rtfRight
Else
TextAllingment = rtfCenter
End If
PutRegString "Textcolor", CStr(TextColor)
PutRegString "TextBold", CStr(TextBold)
PutRegString "TextItalics", CStr(TextItalics)
PutRegString "TextUnderline", CStr(TextUnderline)
PutRegString "TextAlign", CStr(TextAllingment)

'Field
FieldColor = cbField.SelectedColor
FieldBold = CBool(bldField.Value)
FieldItalics = CBool(itField.Value)
FieldUnderline = CBool(uField.Value)
If lftField.Value = True Then
FieldAllingment = rtfLeft
ElseIf rtField.Value = True Then
FieldAllingment = rtfRight
Else
FieldAllingment = rtfCenter
End If
PutRegString "FieldColor", CStr(FieldColor)
PutRegString "FieldBold", CStr(FieldBold)
PutRegString "FieldItalics", CStr(FieldItalics)
PutRegString "FieldUnderline", CStr(FieldUnderline)
PutRegString "FieldAlign", CStr(FieldAllingment)

'HTTP
HttpColor = cbHTTP.SelectedColor
HttpBold = CBool(bldHTTP.Value)
HttpItalics = CBool(itHTTP.Value)
HttpUnderline = CBool(uHTTP.Value)
If lftHTTP.Value = True Then
HttpAllingment = rtfLeft
ElseIf rtHTTP.Value = True Then
HttpAllingment = rtfRight
Else
HttpAllingment = rtfCenter
End If
PutRegString "HTTPColor", CStr(HttpColor)
PutRegString "HTTPbold", CStr(HttpBold)
PutRegString "HTTPItalics", CStr(HttpItalics)
PutRegString "HTTPUnderline", CStr(HttpUnderline)
PutRegString "HTTPalign", CStr(HttpAllingment)

'Comment
CommentColor = cbComment.SelectedColor
CommentBold = CBool(bldComment.Value)
CommentItalics = CBool(itComment.Value)
CommentUnderline = CBool(uComment.Value)
If lftComment.Value = True Then
CommentAllingment = rtfLeft
ElseIf rtComment.Value = True Then
CommentAllingment = rtfRight
Else
CommentAllingment = rtfCenter
End If
PutRegString "CommentColor", CStr(CommentColor)
PutRegString "CommentBold", CStr(CommentBold)
PutRegString "CommentItalics", CStr(CommentItalics)
PutRegString "CommentUnderline", CStr(CommentUnderline)
PutRegString "CommentAlign", CStr(CommentAllingment)

'Back color
rtBackColor = cbBack.SelectedColor
PutRegString "Backcolor", CStr(rtBackColor)
SaveAddressList
Unload Me



End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
If MsgBox("Are you sure you want to restore the default settings for the search box appearence ?" & vbCrLf & "You will loose the current format.", vbYesNo + vbQuestion + vbDefaultButton1, "Restore defaults") = vbNo Then Exit Sub
cbTitle.SelectedColor = vbRed
cbField.SelectedColor = vbBlack
cbText.SelectedColor = vbBlack
cbHTTP.SelectedColor = vbBlue
cbComment.SelectedColor = vbBlack
cbBack.SelectedColor = 12648447
bldTitle.Value = 1
bldComment.Value = 0
bldField.Value = 1
bldHTTP.Value = 1
bldText.Value = 0
itText.Value = 0
itComment.Value = 1
itField.Value = 0
itHTTP.Value = 0
itTitle.Value = 0
uTitle.Value = 1
uComment.Value = 1
uText.Value = 0
uHTTP.Value = 0
uField.Value = 0
lftComment.Value = True
cntTitle.Value = True
lftText.Value = True
lftHTTP.Value = True
lftField.Value = True
check1.Value = 0
Text1.Text = "Select PSC1.* FROM PSC1 order by IndexID ASC,datesumbitted DESC,Accessed Desc"


End Sub

Private Sub Command4_Click()
cd1.DialogTitle = "Open saved filter"
cd1.InitDir = App.path
cd1.Flags = cdlOFNPathMustExist + cdlOFNFileMustExist + cdlOFNHideReadOnly
cd1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
cd1.FilterIndex = 0
cd1.DefaultExt = "txt"
cd1.CancelError = True
On Error GoTo cderr4
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
Text1.Text = aSQl
Else
MsgBox "Not a valid SQL string"
End If
Close #r%
End If
MousePointer = 0

Exit Sub

cderr4:
Close
MousePointer = 0
Exit Sub
End Sub

Private Sub Command5_Click()
X$ = InputBox("Enter as valid eMail address which the programm will consider as an eMail from Planet Source Code" & vbCrLf & vbCrLf & "e.g. MailingList@planet-source-code.com" & vbCrLf & "Attention:" & vbCrLf & "If you enter an eMail that is not valid you may ruin the database!", "Enter a Valid PSC eMail")
If X$ <> "" Then
    If checkIfEmail(X$) = True Then
            If FindStringExact(List1.hwnd, 0, X$ & vbNullString) = -1 Then
            List1.AddItem X$
            Else
            MsgBox "The new eMail address already exists." & vbCrLf & "Adding it just costs in speed.", vbInformation, "Exising eMail"
            End If
    Else
    MsgBox "'" & X$ & "'" & vbCrLf & "Does not seem like a valid eMail address at all.", vbCritical, "eMail Error"
    End If
End If




End Sub

Private Sub Command6_Click()
If List1.ListIndex > -1 And List1.ListIndex < List1.ListCount Then
List1.RemoveItem List1.ListIndex
List1.SetFocus
End If
End Sub

Private Sub Command7_Click()
List1.Clear

End Sub


Private Sub Command8_Click()
frmFind_eMails.Show 1, Me

End Sub

Private Sub Form_Load()
Dim timerobj As New ArielTimer
check1.Value = Abs(LoadSQLatStartup)
Text1.Text = StartUpSQL

cbTitle.SelectedColor = TitleColor
cbText.SelectedColor = TextColor
cbField.SelectedColor = FieldColor
cbHTTP.SelectedColor = HttpColor
cbComment.SelectedColor = CommentColor
cbBack.SelectedColor = rtBackColor

bldTitle.Value = Abs(TitleBold)
bldField.Value = Abs(FieldBold)
bldText.Value = Abs(TextBold)
bldHTTP.Value = Abs(HttpBold)
bldComment.Value = Abs(CommentBold)

itTitle.Value = Abs(TitleItalics)
itField.Value = Abs(FieldItalics)
itText.Value = Abs(TextItalics)
itHTTP.Value = Abs(HttpItalics)
itComment.Value = Abs(CommentItalics)

uTitle.Value = Abs(TitleUnderline)
uField.Value = Abs(FieldUnderline)
uText.Value = Abs(TextUnderline)
uHTTP.Value = Abs(HttpUnderline)
uComment.Value = Abs(CommentUnderline)

If TitleAllingment = rtfLeft Then
    lftTitle.Value = True
ElseIf TitleAllingment = rtfCenter Then
    cntTitle.Value = True
Else
    rtTitle.Value = True
End If

If TextAllingment = rtfLeft Then
    lftText.Value = True
ElseIf TextAllingment = rtfCenter Then
    cntText.Value = True
Else
    rtText.Value = True
End If

If FieldAllingment = rtfLeft Then
    lftField.Value = True
ElseIf FieldAllingment = rtfCenter Then
    cntField.Value = True
Else
    rtField.Value = True
End If

If HttpAllingment = rtfLeft Then
    lftHTTP.Value = True
ElseIf HttpAllingment = rtfCenter Then
    cntHTTP.Value = True
Else
    rtHTTP.Value = True
End If

If CommentAllingment = rtfLeft Then
    lftComment.Value = True
ElseIf CommentAllingment = rtfCenter Then
    cntComment.Value = True
Else
    rtComment.Value = True
End If

LoadAddressList



End Sub


Private Sub itComment_Click()
rt1.SelStart = CommentStart
rt1.SelLength = Commentlength
rt1.SelItalic = CBool(itComment.Value)
rt1.SelLength = 0

End Sub

Private Sub itField_Click()
For i% = 0 To 99 Step 33
rt1.SelStart = 16 + i%
rt1.SelLength = 6
rt1.SelItalic = CBool(itField.Value)
rt1.SelLength = 0
Next i%
rt1.SelStart = 148
rt1.SelLength = 9
rt1.SelItalic = CBool(itField.Value)
rt1.SelLength = 0

rt1.SelStart = 164
rt1.SelLength = 10
rt1.SelItalic = CBool(itField.Value)
rt1.SelLength = 0

rt1.SelStart = 201
rt1.SelLength = 9
rt1.SelItalic = CBool(itField.Value)
rt1.SelLength = 0

End Sub

Private Sub itHTTP_Click()
Dim L As Boolean
L = CBool(itHTTP)
rt1.SelStart = 157
rt1.SelLength = 5
rt1.SelItalic = L

rt1.SelStart = 174
rt1.SelLength = 25
rt1.SelItalic = L

rt1.SelStart = 210
rt1.SelLength = 25
rt1.SelItalic = L

rt1.SelLength = 0

End Sub

Private Sub itText_Click()
rt1.SelStart = 22
rt1.SelLength = 26
rt1.SelItalic = CBool(itText.Value)
rt1.SelLength = 0

rt1.SelStart = 55
rt1.SelLength = 26
rt1.SelItalic = CBool(itText.Value)
rt1.SelLength = 0

rt1.SelStart = 88
rt1.SelLength = 26
rt1.SelItalic = CBool(itText.Value)
rt1.SelLength = 0

rt1.SelStart = 121
rt1.SelLength = 26
rt1.SelItalic = CBool(itText.Value)
rt1.SelLength = 0

End Sub

Private Sub itTitle_Click()
rt1.SelStart = TitleStart
rt1.SelLength = TitleLength
rt1.SelItalic = CBool(itTitle.Value)
rt1.SelLength = 0

End Sub


Private Sub lftComment_Click()
rt1.SelStart = CommentStart
rt1.SelLength = Commentlength
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0

End Sub

Private Sub lftField_Click()
For i% = 0 To 99 Step 33
rt1.SelStart = 16 + i%
rt1.SelLength = 6
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0
Next i%
rt1.SelStart = 148
rt1.SelLength = 9
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0

rt1.SelStart = 164
rt1.SelLength = 10
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0

rt1.SelStart = 201
rt1.SelLength = 9
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0

End Sub

Private Sub lftHTTP_Click()
rt1.SelStart = 157
rt1.SelLength = 5
rt1.SelAlignment = rtleft

rt1.SelStart = 174
rt1.SelLength = 25
rt1.SelAlignment = rtleft

rt1.SelStart = 210
rt1.SelLength = 25
rt1.SelAlignment = rtleft

rt1.SelLength = 0
End Sub

Private Sub lftText_Click()
rt1.SelStart = 22
rt1.SelLength = 26
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0

rt1.SelStart = 55
rt1.SelLength = 26
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0

rt1.SelStart = 88
rt1.SelLength = 26
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0

rt1.SelStart = 121
rt1.SelLength = 26
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0


End Sub

Private Sub lftTitle_Click()
rt1.SelStart = TitleStart
rt1.SelLength = TitleLength
rt1.SelAlignment = rtfLeft
rt1.SelLength = 0

End Sub

Private Sub rtComment_Click()
rt1.SelStart = CommentStart
rt1.SelLength = Commentlength
rt1.SelAlignment = rtfRight
rt1.SelLength = 0

End Sub

Private Sub rtField_Click()
For i% = 0 To 99 Step 33
rt1.SelStart = 16 + i%
rt1.SelLength = 6
rt1.SelAlignment = rtfRight
rt1.SelLength = 0
Next i%
rt1.SelStart = 148
rt1.SelLength = 9
rt1.SelAlignment = rtfRight
rt1.SelLength = 0

rt1.SelStart = 164
rt1.SelLength = 10
rt1.SelAlignment = rtfRight
rt1.SelLength = 0

rt1.SelStart = 201
rt1.SelLength = 9
rt1.SelAlignment = rtfRight
rt1.SelLength = 0

End Sub

Private Sub rtHTTP_Click()
rt1.SelStart = 157
rt1.SelLength = 5
rt1.SelAlignment = rtfRight

rt1.SelStart = 174
rt1.SelLength = 25
rt1.SelAlignment = rtfRight

rt1.SelStart = 210
rt1.SelLength = 25
rt1.SelAlignment = rtfRight

rt1.SelLength = 0

End Sub

Private Sub rtText_Click()
rt1.SelStart = 22
rt1.SelLength = 26
rt1.SelAlignment = rtfRight
rt1.SelLength = 0

rt1.SelStart = 55
rt1.SelLength = 26
rt1.SelAlignment = rtfRight
rt1.SelLength = 0

rt1.SelStart = 88
rt1.SelLength = 26
rt1.SelAlignment = rtfRight
rt1.SelLength = 0

rt1.SelStart = 121
rt1.SelLength = 26
rt1.SelAlignment = rtfRight
rt1.SelLength = 0
End Sub

Private Sub rtTitle_Click()
rt1.SelStart = TitleStart
rt1.SelLength = TitleLength
rt1.SelAlignment = rtfRight
rt1.SelLength = 0

End Sub

Private Sub uComment_Click()
rt1.SelStart = CommentStart
rt1.SelLength = Commentlength
rt1.SelUnderline = CBool(uComment.Value)
rt1.SelLength = 0

End Sub

Private Sub uField_Click()
For i% = 0 To 99 Step 33
rt1.SelStart = 16 + i%
rt1.SelLength = 6
rt1.SelUnderline = CBool(uField.Value)
rt1.SelLength = 0
Next i%
rt1.SelStart = 148
rt1.SelLength = 9
rt1.SelUnderline = CBool(uField.Value)
rt1.SelLength = 0

rt1.SelStart = 164
rt1.SelLength = 10
rt1.SelUnderline = CBool(uField.Value)
rt1.SelLength = 0

rt1.SelStart = 201
rt1.SelLength = 9
rt1.SelUnderline = CBool(uField.Value)
rt1.SelLength = 0

End Sub

Private Sub uHTTP_Click()
Dim L As Boolean
L = CBool(uHTTP)
rt1.SelStart = 157
rt1.SelLength = 5
rt1.SelUnderline = L

rt1.SelStart = 174
rt1.SelLength = 25
rt1.SelUnderline = L

rt1.SelStart = 210
rt1.SelLength = 25
rt1.SelUnderline = L

rt1.SelLength = 0
End Sub

Private Sub uText_Click()
rt1.SelStart = 22
rt1.SelLength = 26
rt1.SelUnderline = CBool(uText.Value)
rt1.SelLength = 0

rt1.SelStart = 55
rt1.SelLength = 26
rt1.SelUnderline = CBool(uText.Value)
rt1.SelLength = 0

rt1.SelStart = 88
rt1.SelLength = 26
rt1.SelUnderline = CBool(uText.Value)
rt1.SelLength = 0

rt1.SelStart = 121
rt1.SelLength = 26
rt1.SelUnderline = CBool(uText.Value)
rt1.SelLength = 0

End Sub

Private Sub uTitle_Click()
rt1.SelStart = TitleStart
rt1.SelLength = TitleLength
rt1.SelUnderline = CBool(uTitle.Value)
rt1.SelLength = 0

End Sub


