VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "Tip of the Day"
   ClientHeight    =   3285
   ClientLeft      =   3990
   ClientTop       =   3180
   ClientWidth     =   5415
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5415
   Begin VB.CheckBox check1 
      Caption         =   "Don't show this again"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2940
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   2715
      Left            =   120
      Picture         =   "frmMessage.frx":0742
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   1785
         Left            =   285
         TabIndex        =   3
         Top             =   420
         Width           =   3105
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If frmMessageCheckKey <> "" Then
SaveSetting "PSC Database", "Settings", frmMessageCheckKey, CStr(CBool(check1.Value))
End If
End Sub

Private Sub cmdOK_Click()
If frmMessageCheckKey <> "" Then
SaveSetting "PSC Database", "Settings", frmMessageCheckKey, CStr(CBool(check1.Value))
End If
Unload Me

End Sub

Private Sub Form_Load()
Me.Caption = frmMessageTitle
cmdOK.Caption = frmMessageButtonCaption
check1.Caption = frmMessageCheckCaption
check1.Visible = frmMessageCheckVisible
Label1.Caption = frmMessageMessage

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMessageTitle = ""
frmMessageButtonCaption = ""
frmMessageCheckCaption = ""
frmMessageCheckVisible = -1
frmMessageMessage = ""
frmMessageCheckKey = ""

End Sub
