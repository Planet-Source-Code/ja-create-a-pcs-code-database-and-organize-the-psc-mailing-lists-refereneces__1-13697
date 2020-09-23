VERSION 5.00
Begin VB.Form frmChange 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   420
      Left            =   3810
      TabIndex        =   6
      Top             =   375
      Width           =   360
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   645
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1920
      TabIndex        =   2
      Top             =   225
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   3090
      TabIndex        =   1
      Top             =   2430
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   525
      Left            =   1800
      TabIndex        =   0
      Top             =   2415
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Datesubmitted year"
      Height          =   210
      Left            =   225
      TabIndex        =   5
      Top             =   735
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Listadate year"
      Height          =   210
      Left            =   255
      TabIndex        =   4
      Top             =   270
      Width           =   1275
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form1.ListdateYeartoChange = Val(Text1.Text)
Form1.DatesubmittedYeartoChange = Val(Text2.Text)
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub


Private Sub Command3_Click()
Text2.Text = Text1.Text
Command1_Click
End Sub

Private Sub Form_Activate()
Text2.SetFocus
End Sub

Private Sub Form_Load()
Text1.Text = Form1.ListdateYeartoChange
Text2.Text = Form1.DatesubmittedYeartoChange


End Sub


