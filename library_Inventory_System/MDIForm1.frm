VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00808080&
   Caption         =   "Library System Version 2.1.0"
   ClientHeight    =   4500
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5445
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4125
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "4:24 AM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/11/2003"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   3  'Align Left
      Height          =   3645
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   6429
      ButtonWidth     =   820
      ButtonHeight    =   794
      Style           =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep11"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep12"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn11"
            Object.ToolTipText     =   "Calculator"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn12"
            Object.ToolTipText     =   "Notepad"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn13"
            Object.ToolTipText     =   "Web Explorer"
            ImageIndex      =   37
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn14"
            Object.ToolTipText     =   "MS Paint"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn15"
            Object.ToolTipText     =   "Window Explorer"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn21"
            Object.ToolTipText     =   "Password Settings"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep15"
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn1"
            Object.ToolTipText     =   "Barrowers"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn2"
            Object.ToolTipText     =   "Books"
            ImageIndex      =   35
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn3"
            Object.ToolTipText     =   "Barrowed Books"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn4"
            Object.ToolTipText     =   "Due Books"
            ImageIndex      =   34
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn5"
            Object.ToolTipText     =   "Returned Books"
            ImageIndex      =   36
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn6"
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn9"
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep2"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn7"
            Object.ToolTipText     =   "ToolBar Align Left"
            ImageIndex      =   38
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn8"
            Object.ToolTipText     =   "ToolBar Align Right"
            ImageIndex      =   39
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":155E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2532
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":36E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":489A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5174
            Key             =   "btn11"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6328
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":68C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6E5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7736
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8010
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":88EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":913E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":96D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A6AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AC46
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B1E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C394
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C92E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":CEC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D7A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E07C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E956
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F230
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":FB0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":103E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11258
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":113B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1194C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11AA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnubarrowers 
         Caption         =   "&Barrowers"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnubooks 
         Caption         =   "B&ooks"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnubarrowed 
         Caption         =   "Ba&rrowed Books"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnudue 
         Caption         =   "&Due Books"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuunreturned 
         Caption         =   "&Returned Books"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnusettings 
         Caption         =   "&Settings"
         Shortcut        =   {F6}
      End
      Begin VB.Menu blnk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnurep 
         Caption         =   "Repor&ts"
         Shortcut        =   {F7}
      End
      Begin VB.Menu blnk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^{F1}
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnuleft 
         Caption         =   "&ToolBar Align Left"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuright 
         Caption         =   "&ToolBar Align Right"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnucalc 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnunote 
         Caption         =   "&Notepad"
      End
      Begin VB.Menu mnuweb 
         Caption         =   "&Web Browser"
      End
      Begin VB.Menu mnupaint 
         Caption         =   "&MS Paint"
      End
      Begin VB.Menu mnuwin 
         Caption         =   "Win&dow Explorer"
      End
      Begin VB.Menu blnk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnupass 
         Caption         =   "&Password Settings"
      End
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "&Window"
      Begin VB.Menu mnucascade 
         Caption         =   "&Arrange Cascade"
      End
      Begin VB.Menu mnuhori 
         Caption         =   "Arrange &Horizontal Tile"
      End
      Begin VB.Menu mnuvert 
         Caption         =   "Arrange &Vertical Tile"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnusys 
         Caption         =   "System &Requirements"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
frmSplash.Show vbModal
DE1.Connection1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
DE1.rsCommand1.Open
DE1.rsCommand2.Open
DE1.rsCommand3.Open
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim reply As Integer
reply = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo, "Library System")
If reply = vbYes Then
    End
Else
    Cancel = 1
End If
End Sub

Private Sub mnuabout_Click()
frmSplash.Show vbModal
End Sub

Private Sub mnubarrowed_Click()
Form1.Show
    Form1.SSTab1.Tab = 2
End Sub

Private Sub mnubarrowers_Click()
Form1.Show
    Form1.SSTab1.Tab = 0
End Sub

Private Sub mnubooks_Click()
Form1.Show
    Form1.SSTab1.Tab = 1
End Sub

Private Sub mnucalc_Click()
On Error Resume Next
Shell "Calc.exe", vbMaximizedFocus
End Sub

Private Sub mnucascade_Click()
Me.Arrange (0)
End Sub

Private Sub mnudue_Click()
Form1.Show
    Form1.SSTab1.Tab = 3
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnuhori_Click()
Me.Arrange (1)
End Sub

Private Sub mnuleft_Click()
Toolbar2.Align = 3
End Sub

Private Sub mnunote_Click()
On Error Resume Next
Shell "Notepad.exe", vbMaximizedFocus
End Sub

Private Sub mnupaint_Click()
On Error Resume Next
Shell "MSpaint.exe", vbMaximizedFocus
End Sub

Private Sub mnupass_Click()
On Error Resume Next
Form8.Show vbModal
End Sub

Private Sub mnurep_Click()
Form10.Show
End Sub

Private Sub mnuright_Click()
Toolbar2.Align = 4
End Sub

Private Sub mnusettings_Click()
Form1.Show
    Form1.SSTab1.Tab = 5
End Sub

Private Sub mnusys_Click()
MsgBox "Processor Speed: 132 Mhz" & vbCrLf & "System Memory: 32MB" & vbCrLf & "Operaing System: Windows 95,98,NT,ME,2000,XP,CE", vbInformation, "Minimum System Requirements"
End Sub


Private Sub mnuunreturned_Click()
Form1.Show
    Form1.SSTab1.Tab = 4
End Sub

Private Sub mnuvert_Click()
Me.Arrange (2)
End Sub

Private Sub mnuweb_Click()
On Error Resume Next
frmBrowser.Show
End Sub

Private Sub mnuwin_Click()
On Error Resume Next
Shell "Explorer.exe", vbMaximizedFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim A As String
A = Button.Key
Select Case A
Case "btn1"
    Form1.Show
    Form1.SSTab1.Tab = 0
Case "btn2"
    Form1.Show
    Form1.SSTab1.Tab = 1
Case "btn3"
    Form1.Show
    Form1.SSTab1.Tab = 2
Case "btn4"
    Form1.Show
    Form1.SSTab1.Tab = 3
Case "btn5"
    Form1.Show
    Form1.SSTab1.Tab = 4
Case "btn6"
    Form1.Show
    Form1.SSTab1.Tab = 5
Case "btn7"
    Toolbar2.Align = 3
Case "btn8"
    Toolbar2.Align = 4
Case "btn9"
    Form10.Show
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim A As String
A = Button.Key
Select Case A
Case "btn11"
    On Error Resume Next
    Shell "Calc.exe", vbMaximizedFocus
Case "btn12"
    On Error Resume Next
    Shell "Notepad.exe", vbMaximizedFocus
Case "btn13"
    On Error Resume Next
    frmBrowser.Show
Case "btn14"
    On Error Resume Next
    Shell "MSpaint.exe", vbMaximizedFocus
Case "btn15"
    On Error Resume Next
    Shell "Explorer.exe", vbMaximizedFocus
Case "btn21"
    On Error Resume Next
    Form8.Show vbModal
End Select
End Sub
