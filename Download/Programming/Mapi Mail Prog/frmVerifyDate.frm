VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerifyDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify listdate"
   ClientHeight    =   3285
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   2655
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   915
      Width           =   1725
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   2655
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1280
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
      Height          =   930
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2250
      Width           =   4155
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   2655
      TabIndex        =   3
      Top             =   1620
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24444929
      CurrentDate     =   36872
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Listdate"
      Height          =   270
      Left            =   210
      TabIndex        =   10
      Top             =   1635
      Width           =   1980
   End
   Begin VB.Label Label4 
      Caption         =   "Message text first lines"
      Height          =   270
      Left            =   225
      TabIndex        =   9
      Top             =   1950
      Width           =   1980
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Message received date:"
      Height          =   270
      Left            =   225
      TabIndex        =   6
      Top             =   1320
      Width           =   1980
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Message ID"
      Height          =   270
      Left            =   225
      TabIndex        =   5
      Top             =   945
      Width           =   1980
   End
   Begin VB.Label Label1 
      Caption         =   $"frmVerifyDate.frx":0000
      Height          =   750
      Left            =   225
      TabIndex        =   2
      Top             =   75
      Width           =   3960
   End
End
Attribute VB_Name = "frmVerifyDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
vCancel = True
Unload Me

End Sub

Private Sub Form_Load()
Text1.Text = modUserInput.vTXT
Text3.Text = vMsgID
Text2.Text = vMsgDate
DTPicker1.Value = vMsgDate


End Sub


Private Sub OKButton_Click()
vDate = DTPicker1.Value
vCancel = False
Unload Me

End Sub

