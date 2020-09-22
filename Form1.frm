VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Test Application"
   ClientHeight    =   4515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ACPRibbon ACPRibbon1 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3069
      BackColor       =   4210752
      ForeColor       =   -2147483630
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":119F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1752
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "For last, insert in Form_Load all code to create Tabs, Categories and Buttons."
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To create the image of the button with transparency, saves as GIF with tranparency."
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   3000
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Insert in ImageList All button images."
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Insert the ACPRibbon and a ImageList to your form."
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#######################################
'#                                     #
'#           ACP Ribbon 2007           #
'#            (Ribbon Only)            #
'#                  by                 #
'#      adrianopaladini@gmail.com      #
'#                                     #
'#                                     #
'#  Visual from Office 2007 Beta 2 TR  #
'#                                     #
'#   Please Don´t Remove Author Info!  #
'#                                     #
'#######################################

'------------------------------------------------
' Bugs:
'
' Please report to:
'
'         adrianopaladini@gmail.com
'
'------------------------------------------------

Dim Theme As Integer

Private Sub Form_Load()

Theme = 0

'# SET Theme
ACPRibbon1.Theme = Theme    ' 0 - Black
                            ' 1 - Blue
                            ' 2 - Silver
                        

'# OPTIONAL - Load Background for Form.
Image1.Left = 0
Image1.Top = ACPRibbon1.Height
Image1.Picture = ACPRibbon1.LoadBackground

'# OPTIONAL - Load Background for Form
Form1.BackColor = ACPRibbon1.BackColor

'# OPTIONAL - Load Forecolor for Labels
Label1.ForeColor = ACPRibbon1.ForeColor
Label2.ForeColor = ACPRibbon1.ForeColor
Label3.ForeColor = ACPRibbon1.ForeColor
Label4.ForeColor = ACPRibbon1.ForeColor


'# Set ImageList to use for icons
ACPRibbon1.ImageList = ImageList1

'# Set Buttons on Center verticaly    (True = Center, False(Default) = Align on Top)
ACPRibbon1.ButtonCenter = False

'# Add Tabs ---   ID - Caption
ACPRibbon1.AddTab "1", "Tab 1"
ACPRibbon1.AddTab "2", "Tab 2"
ACPRibbon1.AddTab "3", "Sample Tab"
ACPRibbon1.AddTab "4", "New Tab"
ACPRibbon1.AddTab "5", "WOW"

'# Add Cats ---   ID - Tab - Caption - ShowDialogButton
ACPRibbon1.AddCat "1", "1", "Group 1", False
ACPRibbon1.AddCat "2", "1", "One very large group", True
ACPRibbon1.AddCat "3", "1", "Test", True
ACPRibbon1.AddCat "4", "2", "More one group", True
ACPRibbon1.AddCat "5", "2", "Hi!", False
ACPRibbon1.AddCat "6", "3", "Hello World!", False

'# Add Button ---    ID - Cat - Capt. - Icons -   More Arrow   - ToolTip
ACPRibbon1.AddButton "0", "1", "CHANGE" & vbNewLine & "THEME", 4
ACPRibbon1.AddButton "1", "1", "Table", 1, False, "Insert a new Table"
ACPRibbon1.AddButton "2", "1", "Insert Picture", 2
ACPRibbon1.AddButton "3", "1", "Insert" & vbNewLine & "Picture", 2
ACPRibbon1.AddButton "4", "2", "Graph", 3
ACPRibbon1.AddButton "5", "2", "Graph", 3, True
ACPRibbon1.AddButton "6", "3", "Clip Art", 4
ACPRibbon1.AddButton "7", "4", "SmartDraw", 5

'# Repaint Ribbon
ACPRibbon1.Refresh

End Sub

Private Sub Form_Resize()

'# this procedure will resize the ribbon
ACPRibbon1.Resize

End Sub

Private Sub ACPRibbon1_TabClick(ByVal ID As String, ByVal Caption As String)

'# This Event occurs when click on each tab
MsgBox "TabSelected: " & ID & "--" & Caption

End Sub

Private Sub ACPRibbon1_CatClick(ByVal ID As String, ByVal Caption As String)

'# This Event occurs when click on each ShowDialogButton for each Categorie
MsgBox "ShowDialogClick: " & ID & "--" & Caption

End Sub

Private Sub ACPRibbon1_ButtonClick(ByVal ID As String, ByVal Caption As String)


'# This Event occurs when click on each Button
MsgBox "ButtonClick: " & ID & "--" & Caption



'# CHANGE THEME
If ID = 0 Then
    Theme = Theme + 1
    If Theme = 3 Then Theme = 0
    '# Set Theme
    ACPRibbon1.Theme = Theme
    '# Refresh control
    ACPRibbon1.Refresh
    
    '# OPTIONAL - Load Background for Form.
    Image1.Picture = ACPRibbon1.LoadBackground
    
    '# OPTIONAL - Load Background for Form
    Form1.BackColor = ACPRibbon1.BackColor
    
    '# OPTIONAL - Load Forecolor for Labels
    Label1.ForeColor = ACPRibbon1.ForeColor
    Label2.ForeColor = ACPRibbon1.ForeColor
    Label3.ForeColor = ACPRibbon1.ForeColor
    Label4.ForeColor = ACPRibbon1.ForeColor
End If


End Sub

