VERSION 5.00
Begin VB.Form StartUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chose example"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "MDI Form"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Standard Form"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "StartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
MDIForm1.Show
End Sub
