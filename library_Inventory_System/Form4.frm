VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1800
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   50266113
      CurrentDate     =   37797
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   50266113
      CurrentDate     =   37797
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   17
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Due :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Barrowed :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barrower's Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:If you want to edit the existing             Barrowed Books, just delete it              and 'Input again'."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3360
      Picture         =   "Form4.frx":06EA
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Picture         =   "Form4.frx":0FB4
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   4080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book NO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Title :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Barrower's ID :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text = "" Then
    Combo1.Text = "1ST YR."
End If
End Sub

Private Sub Command1_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox "All fields required not to be a null value.", vbExclamation, "Library System"
    Exit Sub
End If
With Form1
    .Adodc3.Recordset.AddNew
    .Adodc3.Recordset.Fields(0) = (DataCombo1.Text)
    .Adodc3.Recordset.Fields(1) = (Text2.Text)
    .Adodc3.Recordset.Fields(2) = (Text3.Text)
    .Adodc3.Recordset.Fields(3) = (DataCombo2.Text)
    .Adodc3.Recordset.Fields(4) = (Text4.Text)
    .Adodc3.Recordset.Fields(5) = Format(DTPicker1.Value, "mm/dd/yy")
    .Adodc3.Recordset.Fields(6) = Format(DTPicker2.Value, "mm/dd/yy")
    Call SubtractBookQty
    .Adodc3.Recordset.Update
End With
MsgBox "New Barrowed Book/s has been added." & vbCrLf & "If you want to edit it, just delete it and 'Input again'.", vbInformation, "Library System"
Unload Me
End Sub

Private Sub Command2_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
DataCombo2.Text = ""
DTPicker1.Value = Format(Date, "mm/dd/yy")
DTPicker2.Value = Format(Date, "mm/dd/yy")
End Sub
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub DataCombo1_Change()
If DataCombo1.Text = "" Then
    Text2.Locked = True
    Text3.Locked = True
    DataCombo2.Locked = True
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    DataCombo2.Enabled = False
    DataCombo1.SetFocus
    Command2_Click
    Exit Sub
End If
Form1.Adodc2.Recordset.MoveFirst
Form1.Adodc2.Recordset.Find "BOOK_NO like '" & (DataCombo1.Text) & "'"
If Form1.Adodc2.Recordset.EOF Then
    On Error Resume Next
    Text2.Locked = True
    Text3.Locked = True
    DataCombo2.Locked = True
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
     DataCombo2.Enabled = False
    MsgBox "The Book No. " & (DataCombo1.Text) & " is not exist.Make sure it is correct.", vbExclamation, "Invalid Book No."
    DataCombo1.SetFocus
    Command2_Click
    Exit Sub
Else
On Error Resume Next
    If Val(Form1.Adodc2.Recordset.Fields(9)) <= 0 Then
        MsgBox "All books in titled of " & (Form1.Adodc2.Recordset.Fields(2)) & " has all Barrowed.", vbExclamation, "Library System"
        Exit Sub
    End If
    Text2.Text = (Form1.Adodc2.Recordset.Fields(1))
    Text3.Text = (Form1.Adodc2.Recordset.Fields(2))
    Text2.Locked = False
    Text3.Locked = False
    DataCombo2.Locked = False
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
    DataCombo2.Enabled = True
End If
End Sub

Private Sub DataCombo1_LostFocus()
DataCombo1.Text = UCase(DataCombo1.Text)
End Sub

Private Sub DataCombo2_Change()
If DataCombo2.Text = "" Then
    Text4.Text = ""
    Text4.Locked = True
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    Text4.Locked = True
    Command1.Enabled = False
    DataCombo2.SetFocus
    DTPicker1.Value = Format(Date, "mm/dd/yy")
    DTPicker2.Value = Format(Date, "mm/dd/yy")
    Exit Sub
End If
Form1.Adodc1.Recordset.MoveFirst
Form1.Adodc1.Recordset.Find "BARROWERS_ID like '" & (DataCombo2.Text) & "'"
If Form1.Adodc1.Recordset.EOF Then
    On Error Resume Next
    Text4.Text = ""
    Text4.Locked = True
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    Text4.Locked = True
    Command1.Enabled = False
    MsgBox "The Barrower's ID " & (DataCombo2.Text) & " is not found exist.Make sure it is correct.", vbExclamation, "Invalid Barrower's ID"
    DataCombo2.SetFocus
    DTPicker1.Value = Format(Date, "mm/dd/yy")
    DTPicker2.Value = Format(Date, "mm/dd/yy")
    Exit Sub
Else
    Text4.Text = (Form1.Adodc1.Recordset.Fields(1))
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
    Text4.Locked = False
    Command1.Enabled = True
End If
End Sub

Private Sub Form_Activate()
DataCombo1.SetFocus
End Sub

Private Sub Form_Load()
DTPicker1.Value = Format(Date, "mm/dd/yy")
DTPicker2.Value = Format(Date, "mm/dd/yy")
Call BarrowedkBokNo
Call BarrowedBarID
End Sub
Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
End Sub
Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)
End Sub
Private Sub Text4_LostFocus()
Text4.Text = UCase(Text4.Text)
End Sub
Private Sub Text5_LostFocus()
Text5.Text = UCase(Text5.Text)
End Sub
Private Sub Text6_LostFocus()
Text6.Text = UCase(Text6.Text)
End Sub
Private Sub Text7_LostFocus()
Text7.Text = UCase(Text7.Text)
End Sub
