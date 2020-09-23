VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrintOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   4245
   ClientLeft      =   3285
   ClientTop       =   2445
   ClientWidth     =   6870
   Icon            =   "frmPrintOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5970
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Printer"
      Height          =   1440
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   5265
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1305
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   210
         Width           =   2310
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Properties"
         Height          =   345
         Left            =   3780
         TabIndex        =   10
         Top             =   195
         Width           =   1185
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   210
         Left            =   1305
         TabIndex        =   16
         Top             =   975
         Width           =   2250
      End
      Begin VB.Label Label6 
         Caption         =   "Driver:"
         Height          =   210
         Left            =   180
         TabIndex        =   15
         Top             =   975
         Width           =   1035
      End
      Begin VB.Label Label5 
         Height          =   210
         Left            =   1305
         TabIndex        =   13
         Top             =   615
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "Printer Port:"
         Height          =   210
         Left            =   180
         TabIndex        =   12
         Top             =   615
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Printer Name:"
         Height          =   210
         Left            =   180
         TabIndex        =   11
         Top             =   255
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orientation"
      Height          =   1125
      Left            =   120
      TabIndex        =   6
      Top             =   2835
      Width           =   5265
      Begin VB.OptionButton Option2 
         Caption         =   "Landscape"
         Enabled         =   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   645
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Portrait"
         Enabled         =   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   1185
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1755
         Picture         =   "frmPrintOptions.frx":014A
         Top             =   345
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copies"
      Height          =   945
      Left            =   120
      TabIndex        =   2
      Top             =   1695
      Width           =   5265
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1710
         TabIndex        =   4
         Text            =   "1"
         Top             =   330
         Width           =   540
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2265
         TabIndex        =   3
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   32767
         Min             =   1
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Number of copies"
         Height          =   270
         Left            =   195
         TabIndex        =   5
         Top             =   375
         Width           =   1560
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Print"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2460
      Left            =   0
      ScaleHeight     =   2400
      ScaleWidth      =   3555
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "&Stop"
         Default         =   -1  'True
         Height          =   390
         Left            =   1080
         TabIndex        =   20
         Top             =   1845
         Width           =   1230
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   300
         Left            =   705
         TabIndex        =   19
         Top             =   1290
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   18
         Top             =   540
         Width           =   3450
      End
   End
End
Attribute VB_Name = "frmPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DefaultPrinter As String
Public TotalCopies As Integer
Public CancelPrint As Boolean

Private cSetPrinter As New cSetDfltPrinter

Public StopPrinting As Boolean

Private Sub CancelButton_Click()
Unload Me

End Sub


Private Sub Combo1_Click()
cSetPrinter.SetPrinterAsDefault Combo1.Text
Label5.Caption = Printer.Port
Label7.Caption = Printer.DriverName
If Printer.Orientation = vbPRORPortrait Then
Option1.Value = True
Image1.Picture = LoadResPicture(101, vbResIcon)
Else
Option2.Value = True
Image1.Picture = LoadResPicture(102, vbResIcon)
End If

End Sub

Private Sub Command1_Click()
Dim PrintDef As PRINTER_DEFAULTS
PrintDef.DesiredAccess = PRINTER_ACCESS_USE
Dim g As Long
r% = OpenPrinter(Printer.DeviceName, g&, PrintDef)

f& = PrinterProperties(hwnd, g&)
ClosePrinter g&

Label5.Caption = Printer.Port
Label7.Caption = Printer.DriverName
If Printer.Orientation = vbPRORPortrait Then
Option1.Value = True
Image1.Picture = LoadResPicture(101, vbResIcon)
Else
Option2.Value = True
Image1.Picture = LoadResPicture(102, vbResIcon)
End If

End Sub

Private Sub Command2_Click()
StopPrinting = True

End Sub

Private Sub Form_Load()
DefaultPrinter = Printer.DeviceName

For i% = 0 To Printers.count - 1
Combo1.AddItem Printers(i%).DeviceName
Next i%
For i% = 0 To Combo1.ListCount - 1
If Combo1.List(i%) = DefaultPrinter Then
Combo1.ListIndex = i%
Exit For
End If
Next i%


Label5.Caption = Printer.Port
Label7.Caption = Printer.DriverName
If Printer.Orientation = vbPRORPortrait Then
Option1.Value = True
Image1.Picture = LoadResPicture(101, vbResIcon)
Else
Option2.Value = True
Image1.Picture = LoadResPicture(102, vbResIcon)
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
'set back the default printer as it
'was before the user changed it
If Printer.DeviceName <> DefaultPrinter Then
cSetPrinter.SetPrinterAsDefault DefaultPrinter
End If


End Sub


Private Sub OKButton_Click()

Me.Move Left, Top, 3700, 2820
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Picture1.Visible = True
pb1.Min = 1
pb1.Max = Val(Text1.Text) + 1
DisableX hwnd
Me.Refresh
For i% = 1 To Val(Text1.Text)
DoEvents
If StopPrinting = True Then
Exit For
End If
Label4.Caption = "Printing Copy #" & CStr(i%)
frmSearch.RT1.SelPrint Printer.hDc
pb1.Value = i%
Next i%
Unload Me

End Sub

Private Sub Option2_Click()
Image1.Picture = LoadResPicture(102, vbResIcon)

End Sub


Private Sub Text1_Change()
'UpDown1.Value = CInt(Text1.Text)
'TotalCopies = CInt(Text1.Text)

End Sub


Private Sub UpDown1_Change()
Text1.Text = CStr(UpDown1.Value)

End Sub


