VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify Code reference save"
   ClientHeight    =   2475
   ClientLeft      =   3690
   ClientTop       =   3510
   ClientWidth     =   6030
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Yes for &All"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   690
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Do you want to add the following code reference ?"
      Height          =   2025
      Left            =   30
      TabIndex        =   2
      Top             =   150
      Width           =   4515
      Begin VB.Label Label5 
         Caption         =   "Level"
         Height          =   255
         Left            =   165
         TabIndex        =   7
         Top             =   1605
         Width           =   4170
      End
      Begin VB.Label Label4 
         Caption         =   "Category"
         Height          =   255
         Left            =   165
         TabIndex        =   6
         Top             =   1284
         Width           =   4170
      End
      Begin VB.Label Label3 
         Caption         =   "ListDate"
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   966
         Width           =   4170
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   648
         Width           =   4170
      End
      Begin VB.Label Label1 
         Caption         =   "Title"
         Height          =   255
         Left            =   165
         TabIndex        =   3
         Top             =   330
         Width           =   4170
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&No"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1155
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
pr = No
Unload Me

End Sub

Private Sub Command1_Click()
pr = YesForAll
Unload Me
End Sub


Private Sub Form_Load()
Label1.Caption = "Title: " & CE.Title
Label2.Caption = "Description " & Left$(CE.Text, 42) & "..."
Label3.Caption = "Listdate: " & CE.ListDate
Label4.Caption = "Category: " & CE.Category
Label5.Caption = "Level: " & CE.Level



End Sub


Private Sub OKButton_Click()
pr = yes
Unload Me

End Sub


