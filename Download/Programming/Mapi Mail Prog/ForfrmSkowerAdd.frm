VERSION 5.00
Begin VB.Form ForfrmSkowerAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Slower Add"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   4755
      TabIndex        =   3
      Top             =   3735
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Height          =   3480
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6075
      Begin VB.CommandButton Command1 
         Caption         =   "&Start"
         Height          =   495
         Left            =   4020
         TabIndex        =   1
         Top             =   2115
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   $"ForfrmSkowerAdd.frx":0000
         Height          =   705
         Left            =   195
         TabIndex        =   2
         Top             =   345
         Width           =   5700
      End
   End
End
Attribute VB_Name = "ForfrmSkowerAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub


