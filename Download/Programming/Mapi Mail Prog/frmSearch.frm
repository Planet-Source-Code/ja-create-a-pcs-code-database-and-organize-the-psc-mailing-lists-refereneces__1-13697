VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Code"
   ClientHeight    =   5370
   ClientLeft      =   3555
   ClientTop       =   1020
   ClientWidth     =   8145
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7530
      Top             =   1965
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   210
      Left            =   4410
      TabIndex        =   34
      Top             =   5100
      Visible         =   0   'False
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar st1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   33
      Top             =   5025
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1560
      Index           =   0
      Left            =   135
      ScaleHeight     =   1560
      ScaleWidth      =   6015
      TabIndex        =   5
      Top             =   420
      Width           =   6015
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2505
         TabIndex        =   37
         Text            =   "20"
         Top             =   705
         Width           =   675
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   3180
         TabIndex        =   35
         Top             =   705
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   20
         BuddyControl    =   "Text2"
         BuddyDispid     =   196610
         OrigLeft        =   3270
         OrigTop         =   690
         OrigRight       =   3510
         OrigBottom      =   960
         Increment       =   10
         Max             =   3000
         Min             =   1
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   285
         Width           =   5640
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "matches found"
         Height          =   210
         Left            =   3525
         TabIndex        =   38
         Top             =   705
         Width           =   1785
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Stop Search if there are at least"
         Height          =   240
         Left            =   195
         TabIndex        =   36
         Top             =   705
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Use ""word1 word2 word3"" to search for a pfrase"
         Height          =   240
         Left            =   195
         TabIndex        =   9
         Top             =   1020
         Width           =   5340
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the words  to search  for (Use spaces for separator)"
         Height          =   210
         Left            =   195
         TabIndex        =   8
         Top             =   60
         Width           =   5340
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on 'Options' to define which fields to look in..."
         Height          =   240
         Left            =   195
         TabIndex        =   10
         Top             =   1305
         Width           =   5340
      End
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      Top             =   150
      Visible         =   0   'False
      Width           =   240
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   2490
      Left            =   75
      TabIndex        =   4
      Top             =   2370
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   4392
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmSearch.frx":0442
      MouseIcon       =   "frmSearch.frx":052C
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   6690
      TabIndex        =   3
      Top             =   1515
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&New"
      Height          =   435
      Left            =   6690
      TabIndex        =   2
      Top             =   945
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Find Now"
      Enabled         =   0   'False
      Height          =   435
      Left            =   6690
      TabIndex        =   1
      Top             =   360
      Width           =   1245
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2160
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   3810
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Object.ToolTipText     =   "What to look for"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Object.ToolTipText     =   "Refine your search"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1560
      Index           =   1
      Left            =   180
      ScaleHeight     =   1560
      ScaleWidth      =   6015
      TabIndex        =   6
      Top             =   510
      Width           =   6015
      Begin VB.CheckBox Check6 
         Height          =   150
         Left            =   3030
         TabIndex        =   32
         Top             =   480
         Width           =   165
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Checked"
         Height          =   210
         Left            =   4530
         TabIndex        =   30
         Top             =   645
         Width           =   1350
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Downloaded"
         Height          =   255
         Left            =   4530
         TabIndex        =   29
         Top             =   330
         Width           =   1350
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Refresh Combos"
         Height          =   230
         Left            =   4515
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1080
         Width           =   1450
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3285
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1185
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   3375
         TabIndex        =   20
         Top             =   285
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MgAvantG"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24444929
         CurrentDate     =   36796
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   330
         Width           =   1185
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Note"
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   1005
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Description"
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   660
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Title"
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   330
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   3375
         TabIndex        =   21
         Top             =   585
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MgAvantG"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24444929
         CurrentDate     =   36796
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1185
         Width           =   1185
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Use Defualts"
         Height          =   230
         Left            =   4515
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1335
         Width           =   1450
      End
      Begin VB.Image Image2 
         Height          =   150
         Left            =   3225
         Picture         =   "frmSearch.frx":0846
         Top             =   645
         Width           =   150
      End
      Begin VB.Image Image1 
         Height          =   150
         Left            =   3225
         Picture         =   "frmSearch.frx":0918
         Top             =   330
         Width           =   150
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "|||||||||| Commands |||||||||"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4500
         TabIndex        =   31
         Top             =   870
         Width           =   1480
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "||||||||||| Download |||||||||||"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4500
         TabIndex        =   28
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "|||||||| Compatibility ||||||||"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3000
         TabIndex        =   22
         Top             =   870
         Width           =   1470
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         X1              =   4500
         X2              =   4500
         Y1              =   0
         Y2              =   1600
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         X1              =   4485
         X2              =   4485
         Y1              =   0
         Y2              =   1600
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   2970
         X2              =   2970
         Y1              =   0
         Y2              =   1600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         X1              =   2940
         X2              =   2940
         Y1              =   0
         Y2              =   1600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   1515
         X2              =   1515
         Y1              =   0
         Y2              =   1600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   1485
         X2              =   1485
         Y1              =   0
         Y2              =   1600
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "|||||||||||||||||| Fields ||||||||||||||"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1530
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "||||||||||||||| Date ||||||||||||||||||"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3000
         TabIndex        =   19
         Top             =   0
         Width           =   1470
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "|||||||||| Category |||||||||||"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1530
         TabIndex        =   17
         Top             =   870
         Width           =   1410
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "||||||||||||||| Level ||||||||||||||||"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1515
         TabIndex        =   16
         Top             =   0
         Width           =   1485
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Search Results"
      Height          =   210
      Left            =   285
      TabIndex        =   27
      Top             =   2190
      Width           =   1275
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "S&ave"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActionsCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuActionsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionsOpenURL 
         Caption         =   "Open &URL"
         Enabled         =   0   'False
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuActionsOpenFolder 
         Caption         =   "Open &Folder"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuActionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionsViewDetails 
         Caption         =   "&View details"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmSearch"
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
Private CommentAdded As Boolean



Private CancelSearch As Boolean

Private Const EM_CHARFROMPOS = &HD7


Private Type TPoint
  X As Long
  Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As TPoint) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As TPoint) As Long
Private Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private cur As curs

Private CurRESId As Integer

Private Type ClickTarget
    Type As curs
    Target As String
End Type
Dim cTarget As ClickTarget

Private ToShowWindow As Boolean

Private RunningCurID As Long

Sub AddEntryToRTF(c As Long, f As DAO.Fields)

If CommentAdded = False Then
rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = "Search results for '" & Text1.Text & "'" & vbCrLf & vbCrLf & vbCrLf
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlComment
CommentAdded = True
End If

st1.SimpleText = "Searching..."
rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = CStr(c) & ") " & f("Title").Value & vbNullString & vbCrLf
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlTitle


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = "Category: "
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlField

rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = f("Category").Value & vbNullString & vbCrLf
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stltext

rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = "Level: "
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlField

rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = f("Level").Value & vbNullString & vbCrLf
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stltext


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = "Description: "
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlField

rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
If Len(f("Description").Value & vbNullString) > 60 Then
X$ = Left$(f("Description").Value & vbNullString, 60) & "..." & vbCrLf
Else
X$ = f("Description").Value & vbNullString & vbCrLf
End If
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stltext


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = "Location: "
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlField


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = f("http").Value & vbNullString & vbCrLf
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlhttp


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = "Local Dir: "
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlField


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = f("[LocalDirectory]").Value & vbNullString & vbCrLf
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlhttp


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = "Database IndexID: "
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlField


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = f("[IndexID]").Value & vbNullString & vbCrLf
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlhttp


rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = "Submitted on: "
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlField



rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text) - 1
X$ = Format(f("[DateSumbitted]").Value, "dd/mm/yyyy")
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stltext



rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = " and accessed "
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stlField

rtIndexCount = rtIndexCount + 1
ReDim Preserve rtIndex(rtIndexCount)
rtIndex(rtIndexCount).Start = Len(rt1.Text)
X$ = f("Accessed").Value & " times" & vbCrLf
X$ = X$ & "================================" & vbCrLf & vbCrLf
rt1.Text = rt1.Text & X$
rtIndex(rtIndexCount).End = Len(rt1.Text)
rtIndex(rtIndexCount).Style = stltext



















End Sub


Function DirectFieldSearch(Words() As String, Phrase As Boolean, f1 As Boolean, f2 As Boolean, f3 As Boolean) As Long
Dim Srt As String
Dim iStr As String
Dim cDB As Database
Dim cSet As Recordset
Dim cCounter As Long
cCounter = 0
pb1.Min = 0
'Create database
    
    Set cDB = Workspaces(0).OpenDatabase(App.path & "\psc.mdb", False, True)
    Set cSet = cDB.OpenRecordset(frmCodeEntries.FilterSQL, dbOpenSnapshot)
    If (cSet.BOF And cSet.BOF) Then
    DirectFieldSearch = 0
    cDB.Close
    Exit Function
    End If
cSet.MoveLast
cSet.MoveFirst
pb1.Max = cSet.RecordCount
pb1.Value = pb1.Min
pb1.Visible = True
st1.SimpleText = "Searching..."
If Phrase Then
Srt = UCase(Words(0))
cSet.MoveFirst
rt1.Text = ""
    Do While Not cSet.EOF
        DoEvents
        
        If CancelSearch = True Then
        CancelSearch = False
        Exit Do
        End If
        
        
        If f1 = True Then
        iStr = iStr & " " & cSet.Fields("Title").Value & vbnullsrting
        End If
        If f2 = True Then
        iStr = iStr & " " & cSet.Fields("Description").Value & vbnullsrting
        End If
        If f3 = True Then
        iStr = iStr & " " & cSet.Fields("PersonalNote").Value & vbnullsrting
        End If
        iStr = UCase(iStr)
        If InStr(1, iStr, Srt) > 0 Then
        cCounter = cCounter + 1
        AddEntryToRTF cCounter, cSet.Fields
        If cCounter >= UpDown1.Value Then Exit Do
        End If
        iStr = ""
        cSet.MoveNext
        pb1.Value = pb1.Value + 1
    Loop
DirectFieldSearch = cCounter
cSet.Close
cDB.Close
Exit Function
Else
    Do While Not cSet.EOF
        DoEvents
        
        If CancelSearch = True Then
        CancelSearch = False
        Exit Do
        End If

        If f1 = True Then
        iStr = iStr & " " & cSet.Fields("Title").Value & vbnullsrting
        End If
        If f2 = True Then
        iStr = iStr & " " & cSet.Fields("Description").Value & vbnullsrting
        End If
        If f3 = True Then
        iStr = iStr & " " & cSet.Fields("PersonalNote").Value & vbnullsrting
        End If
        iStr = UCase(iStr)
        For i% = 0 To UBound(Words)
        If InStr(1, iStr, UCase(Words(i%))) > 0 Then
        cCounter = cCounter + 1
        AddEntryToRTF cCounter, cSet.Fields
        If cCounter >= UpDown1.Value Then Exit Do

        End If
        Exit For
        Next i%
        iStr = ""
        cSet.MoveNext
        pb1.Value = pb1.Value + 1
    Loop
DirectFieldSearch = cCounter
cSet.Close
cDB.Close
Exit Function
End If


End Function

Sub EmptyTable(tb As DAO.Recordset)
On Error Resume Next
tb.MoveFirst
Do While Not (tb.EOF And tb.BOF)
tb.MoveFirst
tb.Delete
Loop




End Sub





Sub FormatRTF()
st1.SimpleText = "Formating found entries..."
For i& = 0 To rtIndexCount
SetRTFStyle rt1, rtIndex(i&).Style, rtIndex(i&).Start, rtIndex(i&).End
Next i&
rt1.SelStart = 0
rt1.SelLength = 0

End Sub

Sub InitOptions()
Dim t As Database
Dim S As Recordset
Set t = Workspaces(0).OpenDatabase(App.path & "\PSC.mdb")
Set S = t.OpenRecordset("LevelNames", dbOpenSnapshot)
Combo1.AddItem "<ALL LEVELS>"
Do While Not S.EOF
Combo1.AddItem S.Fields("Levelname").Value & vbNullString
S.MoveNext
Loop
S.Close
Set S = t.OpenRecordset("CategoryNames", dbOpenSnapshot)
Combo2.AddItem "<ALL CATEGORIES>"
Do While Not S.EOF
Combo2.AddItem S.Fields("Categoryname").Value & vbNullString
S.MoveNext
Loop
S.Close
Set S = t.OpenRecordset("CompatibilityNames", dbOpenSnapshot)
Combo3.AddItem "<ALL>"
Do While Not S.EOF
Combo3.AddItem S.Fields("Compatibilityname").Value & vbNullString
S.MoveNext
Loop
S.Close
t.Close



check1.Value = vbChecked
Check2.Value = vbChecked
Check3.Value = vbChecked
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0
DTPicker1.Value = CDate("01/01/" & CStr(Year(Now) - 3))
DTPicker2.Value = Date
ct& = SendMessage(Combo1.hwnd, CB_GETDROPPEDWIDTH, 0, 0)
ret& = SendMessage(Combo1.hwnd, CB_SETDROPPEDWIDTH, 2 * ct&, 0)
ret& = SendMessage(Combo2.hwnd, CB_SETDROPPEDWIDTH, 2.5 * ct&, 0)
ret& = SendMessage(Combo3.hwnd, CB_SETDROPPEDWIDTH, 3 * ct&, 0)
End Sub

Function SearchForStringInDB(SQL As String, Words() As String, Phrase As Boolean, f1 As Boolean, f2 As Boolean, f3 As Boolean) As Long
Dim Srt As String
Dim iStr As String
Dim cDB As Database
Dim cSet As Recordset
Dim cCounter As Long
cCounter = 0

'Create database

    Set cDB = Workspaces(0).OpenDatabase(App.path & "\psc.mdb", False, True)
    Set cSet = cDB.OpenRecordset(SQL, dbOpenSnapshot)
    If (cSet.BOF And cSet.BOF) Then
    SearchForStringInDB = 0
    cDB.Close
    Exit Function
    End If


If Phrase Then
Srt = UCase(Words(0))
cSet.MoveFirst
    Do While Not cSet.EOF
        DoEvents
        
        If CancelSearch = True Then
        CancelSearch = False
        Exit Do
        End If

        If f1 = True Then
        iStr = iStr & " " & cSet.Fields("Title").Value & vbnullsrting
        End If
        If f2 = True Then
        iStr = iStr & " " & cSet.Fields("Description").Value & vbnullsrting
        End If
        If f3 = True Then
        iStr = iStr & " " & cSet.Fields("PersonalNote").Value & vbnullsrting
        End If
        iStr = UCase(iStr)
        If InStr(1, iStr, Srt) > 0 Then
        cCounter = cCounter + 1
        AddEntryToRTF cCounter, cSet.Fields
        If cCounter >= UpDown1.Value Then Exit Do

        End If
        iStr = ""
        cSet.MoveNext
    Loop
SearchForStringInDB = cCounter
cSet.Close
cDB.Close
Exit Function
Else
    Do While Not cSet.EOF
       
        DoEvents
        
        If CancelSearch = True Then
        CancelSearch = False
        Exit Do
        End If
        
        If f1 = True Then
        iStr = iStr & " " & cSet.Fields("Title").Value & vbnullsrting
        End If
        If f2 = True Then
        iStr = iStr & " " & cSet.Fields("Description").Value & vbnullsrting
        End If
        If f3 = True Then
        iStr = iStr & " " & cSet.Fields("PersonalNote").Value & vbnullsrting
        End If
        iStr = UCase(iStr)
        For i% = 0 To UBound(Words)
        If InStr(1, iStr, UCase(Words(i%))) > 0 Then
        cCounter = cCounter + 1
        AddEntryToRTF cCounter, cSet.Fields
        If cCounter >= UpDown1.Value Then Exit Do

        End If
        Exit For
        Next i%
        iStr = ""
        cSet.MoveNext
    Loop
SearchForStringInDB = cCounter
cSet.Close
cDB.Close
Exit Function
End If


End Function

Sub SetRTFStyle(rt As RichTextBox, Style As RTFStyle, cStart As Long, cEnd As Long)
rt.SelStart = cStart
rt.SelLength = cEnd - cStart

Select Case Style
Case 1
    rt.SelAlignment = rtfCenter
    rt.SelFontName = "MS Sans Serif"
    rt.SelFontSize = 10
    rt.SelBold = TitleBold
    rt.SelUnderline = TitleUnderline
    rt.SelColor = TitleColor
    rt.SelItalic = TitleItalics
Case 2
    rt.SelAlignment = FieldAllingment
    rt.SelFontName = "MS Sans Serif"
    rt.SelFontSize = 8
    rt.SelBold = FieldBold
    rt.SelUnderline = FieldUnderline
    rt.SelColor = FieldColor
    rt.SelItalic = FieldItalics
Case 3
    rt.SelAlignment = TextAllingment
    rt.SelFontName = "MS Sans Serif"
    rt.SelFontSize = 8
    rt.SelBold = TextBold
    rt.SelUnderline = TextUnderline
    rt.SelColor = TextColor
    rt.SelItalic = TextItalics
Case 4
    rt.SelAlignment = HttpAllingment
    rt.SelFontName = "MS Sans Serif"
    rt.SelFontSize = 8
    rt.SelBold = HttpBold
    rt.SelUnderline = HttpUnderline
    rt.SelColor = HttpColor
    rt.SelItalic = HttpItalics
Case 5
    rt.SelAlignment = CommentAllingment
    rt.SelFontName = "MS Sans Serif"
    rt.SelFontSize = 10
    rt.SelBold = CommentBold
    rt.SelUnderline = CommentUnderline
    rt.SelColor = CommentColor
    rt.SelItalic = CommentItalics
End Select



End Sub


Private Sub Check1_Click()
If check1.Value = vbUnchecked And Check2.Value = vbUnchecked And Check3.Value = vbUnchecked Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

End Sub


Private Sub Check2_Click()
If check1.Value = vbUnchecked And Check2.Value = vbUnchecked And Check3.Value = vbUnchecked Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

End Sub


Private Sub Check3_Click()
If check1.Value = vbUnchecked And Check2.Value = vbUnchecked And Check3.Value = vbUnchecked Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

End Sub


Private Sub Check6_Click()
DTPicker1.Enabled = Check6.Value = vbChecked
DTPicker2.Enabled = Check6.Value = vbChecked
End Sub


Private Sub Command1_Click()
PutRegString "MaxSearchResults", Text2.Text
CommentAdded = False
If Command1.Caption = "&Cancel" Then
CancelSearch = True
Command1.Caption = "&Find Now"
Me.MousePointer = 0
Else
Command1.Caption = "&Cancel"
MousePointer = 13
rt1.Text = ""
rt1.Refresh
st1.SimpleText = "Initializing search..."
rt1.Visible = False
rtIndexCount = -1

Command2.Enabled = False
Command3.Enabled = False
mnuFile.Enabled = False
mnuActions.Enabled = False
DisableX hwnd
Refresh

Dim Phrase As Boolean
Dim Wordcount As Integer
Dim Combo1SQL As String
Dim Combo2SQL As String
Dim Combo3SQL As String
Dim DateSQL As String
Dim DownloadedSQl As String
Dim CheckedSQL As String
Dim SQL As String



Dim Words() As String
If Left$(Text1.Text, 1) = Chr$(34) Then
    ReDim Words(0)
    If Right$(Text1.Text, 1) = Chr$(34) Then
    Words(0) = Mid$(Text1.Text, 2, Len(Text1.Text) - 2)
    Else
    Words(0) = Right$(Text1.Text, Len(Text1.Text) - 1)
    End If
    Phrase = True
    
Else
    Phrase = False
    Words = Split(Text1.Text, " ")
    Wordcount% = UBound(Words) + 1
End If
'Create SQL
'First we will get all the references that match with the
'Level,Category,Compatibility,Date,Downloaded and Checked Criteria
'Then we will look in the messages if they contain the Words(0)
'1st: Level
st1.SimpleText = "Initializing search: Creating SQL..."
If Combo1.ListIndex = 0 Then
Combo1SQL = ""
Else
Combo1SQL = "Level='" & Combo1.Text & "'"
End If

If Combo2.ListIndex = 0 Then
Combo2SQL = ""
Else
Combo2SQL = "Category='" & Combo2.Text & "'"
End If

If Combo3.ListIndex = 0 Then
Combo3SQL = ""
Else
Combo3SQL = "Compatibility='" & Combo3.Text & "'"
End If

'
'Date options
If Check6.Value = vbChecked Then
    DateSQL = "listDate>=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and listdate<=#" & Format(DTPicker2.Value, "mm/dd/yyyy") & "#"
Else
    DateSQL = ""
End If

'
'Downloaded SQL
If Check4.Value = vbChecked Then
DownloadedSQl = "[HasBeenDownloaded]=true"
Else
DownloadedSQl = ""
End If

'
'Checked SQL
If Check5.Value = vbChecked Then
CheckedSQL = "[HasBeenChecked]=true"
Else
CheckedSQL = ""
End If
If Combo1SQL = "" And Combo2SQL = "" And Combo3SQL = "" And DateSQL = "" And DownloadedSQl = "" And CheckedSQL = "" Then
result = DirectFieldSearch(Words, Phrase, CBool(check1.Value), CBool(Check2.Value), CBool(Check3.Value))
If result > 0 Then
FormatRTF
End If
st1.SimpleText = "Found " & CStr(result) & " code entries."
pb1.Visible = False

Else
Dim SQLStarted As Boolean
SQL = "Select PSC1.* from PSC1 where ("
    If Combo1SQL <> "" Then
    SQL = SQL & Combo1SQL
    SQLStarted = True
    End If
    If Combo2SQL <> "" Then
        If SQLStarted Then
        SQL = SQL & " and " & Combo2SQL
        Else
        SQL = SQL & Combo2SQL
        SQLStarted = True
        End If
    End If
    If Combo3SQL <> "" Then
        If SQLStarted Then
            SQL = SQL & " and " & Combo3SQL
        Else
            SQL = SQL & Combo3SQL
            SQLStarted = True
        End If
    End If
    If DateSQL <> "" Then
        If SQLStarted Then
            SQL = SQL & " and " & DateSQL
        Else
            SQL = SQL & DateSQL
            SQLStarted = True
        End If
    End If
    If DownloadedSQl <> "" Then
        If SQLStarted Then
            SQL = SQL & " and " & DownloadedSQl
        Else
            SQL = SQL & DownloadedSQl
            SQLStarted = True
        End If
    End If
    If CheckedSQL <> "" Then
        If SQLStarted Then
            SQL = SQL & " and " & CheckedSQL
        Else
            SQL = SQL & CheckedSQL
        End If
    End If
    SQL = SQL & ")"
    

    result = SearchForStringInDB(SQL, Words, Phrase, CBool(check1.Value), CBool(Check2.Value), CBool(Check3.Value))
    If result > 0 Then
    FormatRTF
    st1.SimpleText = "Found " & CStr(result) & " code entries."
    pb1.Visible = False
    End If
End If
Command1.Caption = "&Find Now"
Command2.Enabled = True
Command3.Enabled = True
rt1.Visible = True
mnuFile.Enabled = True
mnuActions.Enabled = True
EnableX hwnd
MousePointer = 0

End If

   


End Sub
Private Sub Command2_Click()
Text1.Text = ""
rt1.Text = ""
pb1.Visible = False
st1.SimpleText = ""

End Sub


Private Sub Command3_Click()
Unload Me

End Sub


Private Sub Command4_Click()
check1.Value = vbChecked
Check2.Value = vbChecked
Check3.Value = vbChecked
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0
DTPicker1.Value = 0
DTPicker2.Value = Date
Check4.Value = 0
Check5.Value = 0

End Sub

Private Sub Command5_Click()
MousePointer = 11
'/////////////////////////////////////////////////////////
'Category
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
Combo2.Clear
Combo2.AddItem "<ALL CATEGORIES>"
Set S = t.OpenRecordset("CategoryNames")
EmptyTable S
For i% = 0 To List1.ListCount - 1
S.AddNew
S.Fields("CategoryName").Value = List1.List(i%)
Combo2.AddItem List1.List(i%)
S.Update
Next i%
S.Close
'//////////////////////////////////////////////////////////////
'Level
List1.Clear
Set S = t.OpenRecordset("PSC1", dbOpenSnapshot)
Do While Not S.EOF
    If FindStringExact(List1.hwnd, 0, UCase(S.Fields("Level").Value & vbNullString)) = -1 Then
    List1.AddItem S.Fields("level").Value & vbNullString
    End If
    S.MoveNext
    
Loop
S.Close

Combo1.Clear
Combo1.AddItem "<ALL LEVELS>"
Set S = t.OpenRecordset("LevelNames")
EmptyTable S
For i% = 0 To List1.ListCount - 1
S.AddNew
S.Fields("LevelName").Value = List1.List(i%)
Combo1.AddItem List1.List(i%)
S.Update
Next i%
S.Close
'///////////////////////////////////////////////////////////////////
'Compatibility
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
Combo3.AddItem "<ALL>"
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
MousePointer = 0
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0

End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
InitOptions
Picture1(0).Move TabStrip1.Left + 100, TabStrip1.Top + 400, TabStrip1.Width - 200, TabStrip1.Height - 500
Picture1(1).Move TabStrip1.Left + 100, TabStrip1.Top + 400, TabStrip1.Width - 200, TabStrip1.Height - 500
CurRESId = GetBrowseCursorResID
rt1.BackColor = rtBackColor
Text2.Text = GetRegString("MaxSearchResults", "20")
UpDown1.Value = Val(Text2.Text)
End Sub

Private Sub Form_Resize()
On Error Resume Next
TabStrip1.Width = ScaleWidth - 2000
Picture1(0).Move TabStrip1.Left + 100, TabStrip1.Top + 400, TabStrip1.Width - 200, TabStrip1.Height - 500
Picture1(1).Move TabStrip1.Left + 100, TabStrip1.Top + 400, TabStrip1.Width - 200, TabStrip1.Height - 500
Command1.Left = ScaleWidth - 1700
Command2.Left = ScaleWidth - 1700
Command3.Left = ScaleWidth - 1700
rt1.Height = ScaleHeight - (rt1.Top + 100) - st1.Height
rt1.Width = ScaleWidth - 150
Text1.Width = Picture1(0).ScaleWidth - (Text1.Left + 100)
End Sub


Private Sub mnuActionsCopy_Click()

Clipboard.Clear
If rt1.SelLength > 0 Then
Clipboard.SetText rt1.SelText
Clipboard.SetText rt1.SelRTF, vbCFRTF
Else
Clipboard.SetText rt1.Text
Clipboard.SetText rt1.TextRTF, vbCFRTF
End If
rt1.SetFocus
End Sub

Private Sub mnuFileClose_Click()
Command3_Click
End Sub

Private Sub mnuFileNew_Click()
Command2_Click

End Sub

Private Sub mnuFileOpen_Click()
cd1.DialogTitle = "Open Search results File"
cd1.DefaultExt = "psc"
cd1.Filter = "PSC files (*.PSC)|*.psc|All files (*.*)|*.*"
cd1.FilterIndex = 0
cd1.FileName = ""
cd1.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNHideReadOnly
On Error GoTo FileOpenError
cd1.ShowOpen
If Len(cd1.FileName) > 0 Then
    Open cd1.FileName For Input As #1 Len = 6
    Dim m As String * 6
    Input #1, m$
    Close #1
    If m$ = "@@@@@@" Then
    Open cd1.FileName For Binary As #1
    Put #1, , "{"
    Put #1, , "\"
    Put #1, , "r"
    Put #1, , "t"
    Put #1, , "f"
    Put #1, , "1"
    Close #1
    rt1.LoadFile (cd1.FileName)
    'put back the private stamp
    Open cd1.FileName For Binary As #1
    Put #1, , "@"
    Put #1, , "@"
    Put #1, , "@"
    Put #1, , "@"
    Put #1, , "@"
    Put #1, , "@"
    Close #1
    Else
    MsgBox "Not a valid file", vbCritical, "File Error"
    End If
End If




Exit Sub
FileOpenError:
If Err = cdlCancel Then
Exit Sub
Else
MsgBox Err & " " & Error
Exit Sub
End If
Resume Next




End Sub

Private Sub mnuFilePrint_Click()
On Error Resume Next
selst = rt1.SelStart
selle = rt1.SelLength
rt1.SelLength = 0
frmPrintOptions.Show 1, Me
rt1.SelStart = selst
rt1.SelLength = selle

End Sub


Private Sub mnuFileSave_Click()
cd1.DialogTitle = "Save results file"
cd1.FileName = "PSC Search results for '" & Text1.Text & "'"
cd1.DefaultExt = "psc"
cd1.Filter = "PSC files (*.PSC)|*.psc|All files (*.*)|*.*"
cd1.FilterIndex = 0
cd1.Flags = cdlOFNOverwritePrompt
cd1.InitDir = App.path
On Error GoTo FileSaveError
cd1.ShowSave
If Len(cd1.FileName) > 0 Then
    Open cd1.FileName For Output As #1
    Close #1
    Kill cd1.FileName
    rt1.SaveFile cd1.FileName
    'Put FileStamp
    Open cd1.FileName For Binary As #1
    Put #1, , "@"
    Put #1, , "@"
    Put #1, , "@"
    Put #1, , "@"
    Put #1, , "@"
    Put #1, , "@"
    Close #1
End If
Exit Sub
FileSaveError:
If Err = cdlCancel Then
    Exit Sub
    Resume Next
    Else
    MsgBox Err & " " & Error
    Exit Sub
    Resume Next
End If

    
    

End Sub


Private Sub RT1_Change()
If rt1.SelLength > 0 Or rt1.Text <> "" Then
mnuActionsCopy.Enabled = True
Else
mnuActionsCopy.Enabled = False
End If
' Check For Http
If UCase$(Left$(rt1.SelText, 7)) = "HTTP://" Then
mnuActionsOpenURL.Enabled = True
Else
mnuActionsOpenURL.Enabled = False
End If



End Sub

Private Sub RT1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If rt1.MousePointer = 99 Then
RunningCurID = RunningCurID + 5
rt1.MouseIcon = LoadResPicture(RunningCurID, vbResCursor)
End If

If Button = 2 Then
PopupMenu mnuActions
Exit Sub
End If
If Button = 1 Then
    Select Case cTarget.Type
        Case 4
        Exit Sub
        Case 2
        If Trim(cTarget.Target) <> "" Then
        OpenFolder cTarget.Target
        End If
        Case 1
        If Trim(cTarget.Target) <> "" Then
        openURL cTarget.Target
        End If
        Case 3
        OpenIndex Val(cTarget.Target), frmCodeEntries.Data1
        ToShowWindow = True

    End Select
End If

End Sub


Private Sub RT1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ptMouse As TPoint
' Get screen coordinates
GetCursorPos ptMouse

 ' convert to rich text box coordinates
ScreenToClient rt1.hwnd, ptMouse
Dim MousePos As Long
Dim Charpos As Integer
Dim LineNo As Integer

 MousePos = SendMessageByVal(rt1.hwnd, EM_CHARFROMPOS, CLng(0), ptMouse)
 
 Charpos = LoWord(MousePos)
 LineNo = rt1.GetLineFromChar(Charpos)
 L$ = Mid$(rt1.Text, Charpos + 1, 1)
 If L$ <> vbCr And L$ <> "" Then
    txt$ = GetLine(rt1.hwnd, LineNo)
    'http
    If Left$(txt$, 9) = "Location:" And Charpos > 9 Then
    cTarget.Type = http
    cTarget.Target = ClearUpString(Right$(txt$, Len(txt$) - 10))
    If cur <> http Then
    rt1.MouseIcon = LoadResPicture(CurRESId, vbResCursor)
    RunningCurID = CurRESId
    cur = http
    End If
    rt1.MousePointer = 99
    ElseIf Left$(txt$, 11) = "Local Dir: " And Charpos > 11 Then
    cTarget.Type = folder
    cTarget.Target = ClearUpString(Right$(txt$, Len(txt$) - 11))
    'Folder
    If cur <> folder Then
    rt1.MouseIcon = LoadResPicture(102, vbResCursor)
    RunningCurID = 102
    cur = folder
    End If
    rt1.MousePointer = 99
    'Index
    ElseIf Left$(txt$, 18) = "Database IndexID: " And Charpos > 18 Then
    cTarget.Type = index
    cTarget.Target = ClearUpString(Right$(txt$, Len(txt$) - 18))
    If cur <> index Then
    rt1.MouseIcon = LoadResPicture(101, vbResCursor)
    RunningCurID = 101
    cur = index
    End If
    rt1.MousePointer = 99
    Else
    cTarget.Type = none
    cTarget.Target = ""
    cur = none
    RunningCurID = 101
    rt1.MousePointer = 0
    
    End If
Else
    rt1.MousePointer = 0
End If




End Sub

Private Sub RT1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If ToShowWindow = True Then
ToShowWindow = False
SetWindowPos frmCodeEntries.hwnd, HWND_TOP, 1, 1, 1, 1, SWP_NOMOVE Or SWP_NOSIZE
End If
If rt1.MousePointer = 99 Then
RunningCurID = RunningCurID - 5
rt1.MouseIcon = LoadResPicture(RunningCurID, vbResCursor)
End If

End Sub

Private Sub RT1_SelChange()
If rt1.SelLength > 0 Or rt1.Text <> "" Then
mnuActionsCopy.Enabled = True
Else
mnuActionsCopy.Enabled = False
End If
End Sub


Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.index = 1 Then Command4_Click
Picture1(TabStrip1.SelectedItem.index - 1).ZOrder
If TabStrip1.SelectedItem.index = 1 Then Text1.SetFocus
End Sub


Private Sub Text1_Change()
    Command1.Enabled = Trim(Text1.Text) <> ""
If Command1.Enabled = True Then
Command1.Default = True
Else
Command3.Default = True
End If



End Sub


Private Sub Text2_Change()
UpDown1.Value = Val(Text2.Text)

End Sub

Private Sub UpDown1_Change()
Text2.Text = CStr(UpDown1.Value)

End Sub


