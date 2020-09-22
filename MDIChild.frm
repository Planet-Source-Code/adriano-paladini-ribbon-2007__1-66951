VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ChildMDI 
   Caption         =   "Form2"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   7620
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "ChildMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ProgressBar1.Value = 50
End Sub
