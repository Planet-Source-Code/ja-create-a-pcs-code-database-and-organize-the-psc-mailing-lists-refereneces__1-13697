VERSION 5.00
Begin VB.Form frmViewMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Source"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   420
      Left            =   2745
      TabIndex        =   1
      Top             =   4455
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   4200
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   6345
   End
End
Attribute VB_Name = "frmViewMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Me.Caption = "Message sowrce"
Text1.Text = frmFind_eMails.msgNote_tmp
End Sub


