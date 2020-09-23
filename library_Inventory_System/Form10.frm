VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3465
   Begin TabDlg.SSTab SSTab1 
      Height          =   3900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6879
      _Version        =   393216
      TabOrientation  =   2
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Barrowers"
      TabPicture(0)   =   "Form10.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command23"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "DataCombo1"
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(5)=   "Image1"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Books"
      TabPicture(1)   =   "Form10.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Image2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DataCombo2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Barrowed Books"
      TabPicture(2)   =   "Form10.frx":05C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Image3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton Command2 
         Caption         =   "&Print Preview"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   600
         TabIndex        =   16
         Top             =   2040
         Width           =   2625
      End
      Begin VB.Frame Frame2 
         Caption         =   "Print Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74520
         TabIndex        =   10
         Top             =   1080
         Width           =   2655
         Begin VB.OptionButton Option4 
            Caption         =   "Print by Author"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Print all Books"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print Preview"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -74520
         TabIndex        =   8
         Top             =   3240
         Width           =   2625
      End
      Begin VB.CommandButton Command23 
         Caption         =   "&Print Preview"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -74520
         TabIndex        =   5
         Top             =   3240
         Width           =   2625
      End
      Begin VB.Frame Frame1 
         Caption         =   "Print Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74520
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
         Begin VB.OptionButton Option2 
            Caption         =   "Print all Barrowers"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Print by  Year"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   2415
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   -74520
         TabIndex        =   6
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   -74520
         TabIndex        =   13
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   2400
         Picture         =   "Form10.frx":05DE
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "PRINT BARROWED BOOKS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   -72720
         Picture         =   "Form10.frx":0B68
         Stretch         =   -1  'True
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "PRINT BOOKS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   7
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "PRINT BARROWERS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   -72720
         Picture         =   "Form10.frx":10F2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   840
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DE1.rsCommand2.Filter = ""
If Option3.Value = True Then
    DE1.rsCommand2.Filter = ""
End If
If Option4.Value = True Then
    If DataCombo2.Text = "" Then
        MsgBox "Pls. select a valid Author.", vbExclamation, "Library System"
    Exit Sub
    End If
    DE1.rsCommand2.Filter = "AUTHOR ='" & (DataCombo2.Text) & "'"
End If
DE1.rsCommand2.Close
Rpt2.Show
End Sub

Private Sub Command2_Click()
DE1.rsCommand3.Close
Rpt3.Show
End Sub

Private Sub Command23_Click()
DE1.rsCommand1.Filter = ""
If Option2.Value = True Then
    DE1.rsCommand1.Filter = ""
End If
If Option1.Value = True Then
    If DataCombo1.Text = "" Then
        MsgBox "Pls. select a valid year for barrowers.", vbExclamation, "Library System"
    Exit Sub
    End If
    DE1.rsCommand1.Filter = "CURRENT_YEAR ='" & (DataCombo1.Text) & "'"
End If
DE1.rsCommand1.Close
Rpt1.Show
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
SSTab1.Tab = 0
Option1.Value = True
Option4.Value = True
Call yearPrint
authorPrint
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Label2.Visible = True
    DataCombo1.Visible = True
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    Label2.Visible = False
    DataCombo1.Visible = False
End If
End Sub

Private Sub Option3_Click()
If Option4.Value = False Then
    Label4.Visible = False
    DataCombo2.Visible = False
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
    Label4.Visible = True
    DataCombo2.Visible = True
End If
End Sub

