VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9285
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin Project1.ACPRibbon ACPRibbon1 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   3069
      BackColor       =   4210752
      ForeColor       =   -2147483630
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   2520
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
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":05B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0BD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":119F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1752
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Theme As Integer
Dim fchild As ChildMDI


Private Sub ACPRibbon1_ButtonClick(ByVal ID As String, ByVal Caption As String)
If ID = 0 Then
    Theme = Theme + 1
    If Theme = 3 Then Theme = 0
    '# Set Theme
    ACPRibbon1.Theme = Theme
    '# Refresh control
    ACPRibbon1.Refresh
    
    '# OPTIONAL - Load Background for Form.
    MDIForm1.Picture = ACPRibbon1.LoadBackground
    
    '# OPTIONAL - Load Background for Form
    MDIForm1.BackColor = ACPRibbon1.BackColor
    
    
    '# Search for all MDIChild loaded
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = "ChildMDI" Then
            '# Change Theme from MDIChild Forms
            Forms(i).Picture = ACPRibbon1.LoadBackground
            Forms(i).BackColor = ACPRibbon1.BackColor
            '# Change Forecolor from all Labels on MDIChild forms
            For Each ctl In Forms(i)
                If TypeOf ctl Is Label Then ctl.ForeColor = ACPRibbon1.ForeColor
            Next
        End If
    Next
    
    
    
    
End If

If ID = 1 Then
    '# Open a new Child Form
    Set fchild = New ChildMDI
    fchild.Show
    
    '# Set Theme for new Child Form
    fchild.Picture = ACPRibbon1.LoadBackground
    fchild.BackColor = ACPRibbon1.BackColor
    
End If


End Sub


Private Sub MDIForm_Load()


Theme = 1

'# SET Theme
ACPRibbon1.Theme = Theme    ' 0 - Black
                            ' 1 - Blue
                            ' 2 - Silver
                        

'# OPTIONAL - Load Background for Form.
MDIForm1.Picture = ACPRibbon1.LoadBackground

'# OPTIONAL - Load Background for Form
MDIForm1.BackColor = ACPRibbon1.BackColor

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
ACPRibbon1.AddButton "1", "1", "OPEN CHILD", 1, False, "Open a new form child"
ACPRibbon1.AddButton "2", "1", "Insert Picture", 2
ACPRibbon1.AddButton "3", "1", "Insert" & vbNewLine & "Picture", 2
ACPRibbon1.AddButton "4", "2", "Graph", 3
ACPRibbon1.AddButton "5", "2", "Graph", 3, True
ACPRibbon1.AddButton "6", "3", "Clip Art", 4
ACPRibbon1.AddButton "7", "4", "SmartDraw", 5

'# Repaint Ribbon
ACPRibbon1.Refresh




End Sub
