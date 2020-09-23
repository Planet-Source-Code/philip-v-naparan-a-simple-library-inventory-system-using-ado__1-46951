VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Records"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8040
   Begin MSDataGridLib.DataGrid DataGrid6 
      Height          =   135
      Left            =   1320
      TabIndex        =   125
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   238
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
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
      Left            =   6650
      TabIndex        =   13
      Top             =   5880
      Width           =   1300
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
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
      TabPicture(0)   =   "Form1.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Picture4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Books"
      TabPicture(1)   =   "Form1.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Picture5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Barrowed Books"
      TabPicture(2)   =   "Form1.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command19"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Picture6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command13"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command12"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command11"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command10"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame4"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Due Books"
      TabPicture(3)   =   "Form1.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command14"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "DTPicker1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Picture7"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label11"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Returned Books"
      TabPicture(4)   =   "Form1.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command17"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command16"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Picture8"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame5"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Settings"
      TabPicture(5)   =   "Form1.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame6"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame7 
         Caption         =   "Change Fines"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   117
         Top             =   600
         Width           =   7575
         Begin VB.CommandButton Command21 
            Caption         =   "&Change"
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
            Left            =   6000
            TabIndex        =   118
            Top             =   840
            Width           =   1300
         End
         Begin VB.Label Label50 
            Caption         =   $"Form1.frx":0972
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   126
            Top             =   600
            Width           =   5775
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Add Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         TabIndex        =   116
         Top             =   2400
         Width           =   7575
         Begin VB.TextBox Text19 
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
            Left            =   2640
            TabIndex        =   123
            Top             =   840
            Width           =   4815
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   350
            Left            =   10
            ScaleHeight     =   345
            ScaleWidth      =   2985
            TabIndex        =   122
            Top             =   1680
            Width           =   2980
            Begin MSAdodcLib.Adodc Adodc6 
               Height          =   345
               Left            =   0
               Top             =   0
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   609
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "Adodc1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
         End
         Begin VB.CommandButton Command25 
            Caption         =   "&Refresh"
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
            Left            =   6000
            TabIndex        =   121
            Top             =   1680
            Width           =   1425
         End
         Begin VB.CommandButton Command24 
            Caption         =   "&Delete"
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
            Left            =   4440
            TabIndex        =   120
            Top             =   1680
            Width           =   1425
         End
         Begin VB.CommandButton Command23 
            Caption         =   "&Add"
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
            Left            =   3000
            TabIndex        =   119
            Top             =   1680
            Width           =   1305
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Category :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   124
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command19 
         Caption         =   "&Return"
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
         Left            =   -72960
         TabIndex        =   115
         Top             =   4380
         Width           =   1305
      End
      Begin VB.CommandButton Command17 
         Caption         =   "&Search"
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
         Left            =   -70080
         TabIndex        =   114
         Top             =   4440
         Width           =   1300
      End
      Begin VB.CommandButton Command16 
         Caption         =   "&Refresh"
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
         Left            =   -68640
         TabIndex        =   113
         Top             =   4440
         Width           =   1300
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00808080&
         Height          =   1575
         Left            =   -74880
         ScaleHeight     =   1515
         ScaleWidth      =   7515
         TabIndex        =   108
         Top             =   2760
         Width           =   7575
         Begin MSAdodcLib.Adodc Adodc5 
            Height          =   345
            Left            =   600
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid5 
            Height          =   1215
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   2143
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   " Record:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   112
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "  of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   111
            Top             =   1275
            Width           =   255
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3120
            TabIndex        =   110
            Top             =   1275
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command14 
         Caption         =   "&Refresh"
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
         Left            =   -68640
         TabIndex        =   107
         Top             =   4320
         Width           =   1300
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   -73800
         TabIndex        =   101
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         _Version        =   393216
         Format          =   19136513
         CurrentDate     =   37798
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00808080&
         Height          =   3135
         Left            =   -74880
         ScaleHeight     =   3075
         ScaleWidth      =   7515
         TabIndex        =   100
         Top             =   1080
         Width           =   7575
         Begin MSDataGridLib.DataGrid DataGrid4 
            Height          =   2785
            Left            =   0
            TabIndex        =   103
            Top             =   0
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   4921
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   345
            Left            =   600
            Top             =   2760
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   " Record:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   106
            Top             =   2835
            Width           =   735
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "  of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   105
            Top             =   2835
            Width           =   255
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3120
            TabIndex        =   104
            Top             =   2835
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Book Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         TabIndex        =   79
         Top             =   600
         Width           =   7575
         Begin VB.TextBox Text33 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   89
            Text            =   "Text1"
            Top             =   480
            Width           =   2100
         End
         Begin VB.TextBox Text32 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   88
            Text            =   "Text1"
            Top             =   720
            Width           =   2100
         End
         Begin VB.TextBox Text31 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   960
            Width           =   2625
         End
         Begin VB.TextBox Text30 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   86
            Text            =   "Text1"
            Top             =   1200
            Width           =   2625
         End
         Begin VB.TextBox Text29 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   85
            Text            =   "Text1"
            Top             =   1440
            Width           =   2100
         End
         Begin VB.TextBox Text28 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   84
            Text            =   "Text1"
            Top             =   1680
            Width           =   2100
         End
         Begin VB.TextBox Text27 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   83
            Text            =   "Text1"
            Top             =   480
            Width           =   1860
         End
         Begin VB.TextBox Text26 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   82
            Text            =   "Text1"
            Top             =   720
            Width           =   1860
         End
         Begin VB.TextBox Text25 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   81
            Text            =   "Text1"
            Top             =   960
            Width           =   1860
         End
         Begin VB.TextBox Text24 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   80
            Text            =   "Text1"
            Top             =   1320
            Width           =   1860
         End
         Begin VB.Label Label48 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's ID : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label47 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book Title : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label46 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label45 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book NO : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label44 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Due : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   95
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label43 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Barrowed : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label42 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's Name : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label41 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "No of Days After Due : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   92
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label40 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Returned : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   91
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label38 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Fines : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   90
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00808080&
         Height          =   1575
         Left            =   -74880
         ScaleHeight     =   1515
         ScaleWidth      =   7515
         TabIndex        =   74
         Top             =   2760
         Width           =   7575
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   345
            Left            =   600
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   1215
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   2143
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3120
            TabIndex        =   78
            Top             =   1275
            Width           =   1695
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "  of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   77
            Top             =   1275
            Width           =   255
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   " Record:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   76
            Top             =   1275
            Width           =   735
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "&Refresh"
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
         Left            =   -68640
         TabIndex        =   73
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Delete"
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
         Left            =   -70080
         TabIndex        =   72
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Search"
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
         Left            =   -71520
         TabIndex        =   71
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Add"
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
         Left            =   -74400
         TabIndex        =   70
         Top             =   4380
         Width           =   1305
      End
      Begin VB.Frame Frame4 
         Caption         =   "Book Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         TabIndex        =   55
         Top             =   600
         Width           =   7575
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   480
            Width           =   1980
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "Text1"
            Top             =   1680
            Width           =   2100
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   1440
            Width           =   2100
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   "Text1"
            Top             =   1200
            Width           =   2625
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   58
            Text            =   "Text1"
            Top             =   960
            Width           =   2625
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "Text1"
            Top             =   720
            Width           =   2100
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "Text1"
            Top             =   480
            Width           =   2100
         End
         Begin VB.Label Label37 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's Name : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label36 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Barrowed : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label35 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Due : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   67
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label34 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book NO : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label33 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label32 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book Title : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label31 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's ID : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1200
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Add/Edit"
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
         Left            =   -72960
         TabIndex        =   5
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Search"
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
         Left            =   -71520
         TabIndex        =   6
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Delete"
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
         Left            =   -70080
         TabIndex        =   7
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Refresh"
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
         Left            =   -68640
         TabIndex        =   8
         Top             =   4380
         Width           =   1300
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00808080&
         Height          =   1575
         Left            =   -74880
         ScaleHeight     =   1515
         ScaleWidth      =   7515
         TabIndex        =   43
         Top             =   2760
         Width           =   7575
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   345
            Left            =   600
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   1215
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   2143
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   " Record:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   47
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "  of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   46
            Top             =   1275
            Width           =   255
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3120
            TabIndex        =   45
            Top             =   1275
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Book Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         TabIndex        =   28
         Top             =   600
         Width           =   7575
         Begin VB.Frame Frame3 
            Caption         =   "Book Details"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   4200
            TabIndex        =   48
            Top             =   840
            Width           =   3255
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Remaining :"
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
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Barrowed :"
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
               Left            =   120
               TabIndex        =   53
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity :"
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
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Label21"
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
               Left            =   1080
               TabIndex        =   51
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Label21"
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
               Left            =   1080
               TabIndex        =   50
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Label21"
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
               Left            =   1080
               TabIndex        =   49
               Top             =   720
               Width           =   2055
            End
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   480
            Width           =   2100
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   720
            Width           =   2100
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   960
            Width           =   2745
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   1200
            Width           =   2745
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   1440
            Width           =   2100
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   1680
            Width           =   2100
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   480
            Width           =   2580
         End
         Begin VB.Label Label13 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Author : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label14 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book Title : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label15 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label16 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Book NO : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Price : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   38
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label17 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Year Published : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label18 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Category : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1440
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00808080&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   7515
         TabIndex        =   15
         Top             =   2760
         Width           =   7575
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   345
            Left            =   600
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1215
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   2143
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3120
            TabIndex        =   19
            Top             =   1275
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "  of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   18
            Top             =   1275
            Width           =   255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   " Record:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   1275
            Width           =   735
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Refresh"
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
         Left            =   6360
         TabIndex        =   4
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
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
         Left            =   4920
         TabIndex        =   3
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Search"
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
         Left            =   3480
         TabIndex        =   2
         Top             =   4380
         Width           =   1300
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add/Edit"
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
         Left            =   2040
         TabIndex        =   1
         Top             =   4380
         Width           =   1300
      End
      Begin VB.Frame Frame1 
         Caption         =   "Barrower Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   7575
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   480
            Width           =   2815
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   840
            Width           =   2815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1200
            Width           =   2815
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   255
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   1560
            Width           =   2815
         End
         Begin VB.Image Image1 
            Height          =   1560
            Left            =   5400
            Picture         =   "Form1.frx":09FC
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1800
         End
         Begin VB.Label Label7 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's ID : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Barrower's Name : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label5 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Course : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label6 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Year : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1560
            Width           =   1935
         End
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Date : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   102
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8055
      TabIndex        =   10
      Top             =   0
      Width           =   8055
      Begin VB.Image Image3 
         Height          =   480
         Left            =   7320
         Picture         =   "Form1.frx":0F86
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "This form contains all information on a library."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8055
      TabIndex        =   9
      Top             =   630
      Width           =   8055
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   20
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   645
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc1.Recordset.RecordCount) <= 0 Then
    Adodc1.Caption = "0"
    Label3.Caption = "0"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
Else
    Adodc1.Caption = (Adodc1.Recordset.AbsolutePosition)
    Label3.Caption = (Adodc1.Recordset.RecordCount)
    Text1.Text = (Adodc1.Recordset.Fields(0))
    Text2.Text = (Adodc1.Recordset.Fields(1))
    Text3.Text = (Adodc1.Recordset.Fields(2))
    Text4.Text = (Adodc1.Recordset.Fields(3))
End If
End Sub

Private Sub Adodc2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc2.Recordset.RecordCount) <= 0 Then
    Adodc2.Caption = "0"
    Label26.Caption = "0"
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Label20.Caption = ""
    Label22.Caption = ""
    Label23.Caption = ""
Else
    Adodc2.Caption = (Adodc2.Recordset.AbsolutePosition)
    Label26.Caption = (Adodc2.Recordset.RecordCount)
    Text5.Text = (Adodc2.Recordset.Fields(0))
    Text6.Text = (Adodc2.Recordset.Fields(1))
    Text7.Text = (Adodc2.Recordset.Fields(2))
    Text8.Text = (Adodc2.Recordset.Fields(3))
    Text9.Text = (Adodc2.Recordset.Fields(4))
    Text10.Text = (Adodc2.Recordset.Fields(5))
    Text11.Text = (Adodc2.Recordset.Fields(6))
    Label20.Caption = (Adodc2.Recordset.Fields(7))
    Label22.Caption = (Adodc2.Recordset.Fields(8))
    Label23.Caption = (Adodc2.Recordset.Fields(9))
End If
End Sub

Private Sub Adodc3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc3.Recordset.RecordCount) <= 0 Then
    Adodc3.Caption = "0"
    Label30.Caption = "0"
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
    Text22.Text = ""
Else
    Adodc3.Caption = (Adodc3.Recordset.AbsolutePosition)
    Label30.Caption = (Adodc3.Recordset.RecordCount)
    Text12.Text = (Adodc3.Recordset.Fields(0))
    Text13.Text = (Adodc3.Recordset.Fields(1))
    Text14.Text = (Adodc3.Recordset.Fields(2))
    Text15.Text = (Adodc3.Recordset.Fields(3))
    Text16.Text = (Adodc3.Recordset.Fields(4))
    Text17.Text = (Adodc3.Recordset.Fields(5))
    Text18.Text = (Adodc3.Recordset.Fields(6))
    Text19.Text = (Adodc3.Recordset.Fields(7))
    Text20.Text = (Adodc3.Recordset.Fields(8))
    Text21.Text = (Adodc3.Recordset.Fields(9))
    Text22.Text = (Adodc3.Recordset.Fields(10))
End If
End Sub

Private Sub Adodc4_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc4.Recordset.RecordCount) <= 0 Then
    Adodc4.Caption = "0"
    Label21.Caption = "0"
Else
    Adodc4.Caption = (Adodc4.Recordset.AbsolutePosition)
    Label21.Caption = (Adodc4.Recordset.RecordCount)
End If
End Sub

Private Sub Adodc5_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc5.Recordset.RecordCount) <= 0 Then
    Adodc5.Caption = "0"
    Label28.Caption = "0"
    Text23.Text = ""
    Text24.Text = ""
    Text25.Text = ""
    Text26.Text = ""
    Text27.Text = ""
    Text28.Text = ""
    Text29.Text = ""
    Text30.Text = ""
    Text31.Text = ""
    Text32.Text = ""
    Text33.Text = ""
Else
    Adodc5.Caption = (Adodc5.Recordset.AbsolutePosition)
    Label28.Caption = (Adodc5.Recordset.RecordCount)
    Text33.Text = (Adodc5.Recordset.Fields(0))
    Text32.Text = (Adodc5.Recordset.Fields(1))
    Text31.Text = (Adodc5.Recordset.Fields(2))
    Text30.Text = (Adodc5.Recordset.Fields(3))
    Text29.Text = (Adodc5.Recordset.Fields(4))
    Text28.Text = (Adodc5.Recordset.Fields(5))
    Text27.Text = (Adodc5.Recordset.Fields(6))
    Text26.Text = (Adodc5.Recordset.Fields(7))
    Text25.Text = (Adodc5.Recordset.Fields(8))
    Text24.Text = (Adodc5.Recordset.Fields(9))
    Text23.Text = (Adodc5.Recordset.Fields(10))
End If
End Sub

Private Sub Adodc6_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
If Val(Adodc6.Recordset.RecordCount) <= 0 Then
    Adodc6.Caption = "0"
    Text19.Text = ""
Else
    Adodc6.Caption = (Adodc6.Recordset.AbsolutePosition)
    Text19.Text = (Adodc6.Recordset.Fields(0))
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command10_Click()
Form4.Show vbModal
End Sub

Private Sub Command11_Click()
Dim str1, str2, str3 As String
str1 = InputBox("1. Search by Book No." & vbCrLf & "2. Search by Book Title.", "Search Option")
If str1 = "" Then Exit Sub
If str1 = 1 Then
    str2 = InputBox("Enter BOOK NO  :", "Search by Book No")
    Adodc3.Recordset.Filter = "BOOK_NO ='" & str2 & "'"
Else
    If str1 = 2 Then
    str3 = InputBox("Enter BOOK NAME :", "Search by Book Title")
    Adodc3.Recordset.Filter = "BOOK_TITLE ='" & str3 & "'"
    End If
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
Dim repp2 As Integer
If Val(Adodc3.Recordset.RecordCount) <= 0 Then
    MsgBox "No more Records to be deleted.", vbInformation, "Confirm"
    Exit Sub
End If
repp2 = MsgBox("You are about to delete 1 record." & vbCrLf & "If you click YES, you won't be able to undo this delete operation." & vbCrLf & "Are you sure you want to delete these record?", vbCritical + vbYesNo, "Confirm Delete")
If repp2 = vbYes Then
    Adodc3.Recordset.Delete
    Adodc3.Recordset.MoveNext
    If Adodc3.Recordset.EOF Then
        Adodc3.Recordset.MoveLast
    End If
    Call ReturnBookQty
    MsgBox "Record has been successfuly deleted.", vbInformation, "Confirm"
End If
End Sub

Private Sub Command13_Click()
Adodc3.Refresh
End Sub

Private Sub Command14_Click()
Adodc4.Refresh
Adodc4.Recordset.Filter = "DATE_DUE ='" & Format(DTPicker1.Value, "mm/dd/yy") & "'"
End Sub


Private Sub Command16_Click()
Adodc5.Refresh
End Sub

Private Sub Command17_Click()
Dim str1, str2, str3 As String
str1 = InputBox("1. Search by Book No." & vbCrLf & "2. Search by Book Title.", "Search Option")
If str1 = "" Then Exit Sub
If str1 = 1 Then
    str2 = InputBox("Enter BOOK NO  :", "Search by Book No")
    Adodc5.Recordset.Filter = "BOOK_NO ='" & str2 & "'"
Else
    If str1 = 2 Then
    str3 = InputBox("Enter BOOK NAME :", "Search by Book Title")
    Adodc5.Recordset.Filter = "BOOK_TITLE ='" & str3 & "'"
    End If
End If
End Sub

Private Sub Command19_Click()
Form5.Show vbModal
End Sub

Private Sub Command2_Click()
Form2.Show vbModal
End Sub

Private Sub Command21_Click()
Form6.Show vbModal
End Sub

Private Sub Command22_Click()

End Sub

Private Sub Command23_Click()
Form7.Show vbModal
End Sub

Private Sub Command24_Click()
On Error Resume Next
Dim repp3 As Integer
If Val(Adodc6.Recordset.RecordCount) <= 0 Then
    MsgBox "No more Records to be deleted.", vbInformation, "Confirm"
    Exit Sub
End If
repp3 = MsgBox("You are about to delete 1 record." & vbCrLf & "If you click YES, you won't be able to undo this delete operation." & vbCrLf & "Are you sure you want to delete these record?", vbCritical + vbYesNo, "Confirm Delete")
If repp3 = vbYes Then
    Adodc6.Recordset.Delete
    Adodc6.Recordset.MoveNext
    If Adodc6.Recordset.EOF Then
        Adodc6.Recordset.MoveLast
    End If
    Call ReturnBookQty
    MsgBox "Record has been successfuly deleted.", vbInformation, "Confirm"
End If
End Sub

Private Sub Command25_Click()
Adodc6.Refresh
End Sub

Private Sub Command3_Click()
Dim str1, str2, str3 As String
str1 = InputBox("1. Search by Barrower ID." & vbCrLf & "2. Search by Barrower Name.", "Search Option")
If str1 = "" Then Exit Sub
If str1 = 1 Then
    str2 = InputBox("Enter BARROWER NAME :", "Search by Barrower ID")
    Adodc1.Recordset.Filter = "BARROWERS_ID ='" & str2 & "'"
Else
    If str1 = 2 Then
    str3 = InputBox("Enter BARROWERS ID :", "Search by Barrower Name")
    Adodc1.Recordset.Filter = "BARROWERS_NAME ='" & str3 & "'"
    End If
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim repp As Integer
If Val(Adodc1.Recordset.RecordCount) <= 0 Then
    MsgBox "No more Records to be deleted.", vbInformation, "Confirm"
    Exit Sub
End If
repp = MsgBox("You are about to delete 1 record." & vbCrLf & "If you click YES, you won't be able to undo this delete operation." & vbCrLf & "Are you sure you want to delete these record?", vbCritical + vbYesNo, "Confirm Delete")
If repp = vbYes Then
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveLast
    End If
    MsgBox "Record has been successfuly deleted.", vbInformation, "Confirm"
End If
End Sub

Private Sub Command5_Click()
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Form3.Show vbModal
End Sub

Private Sub Command7_Click()
Dim str1, str2, str3, str4 As String
str1 = InputBox("1. Search by Book No." & vbCrLf & "2. Search by Book Title." & vbCrLf & "3. Search by Author.", "Search Option")
If str1 = "" Then Exit Sub
If str1 = 1 Then
    str2 = InputBox("Enter BOOK NO  :", "Search by Book No")
    Adodc2.Recordset.Filter = "BOOK_NO ='" & str2 & "'"
Else
    If str1 = 2 Then
    str3 = InputBox("Enter BOOK NAME :", "Search by Book Title")
    Adodc2.Recordset.Filter = "BOOK_TITLE ='" & str3 & "'"
    End If
End If
If str1 = 3 Then
    str4 = InputBox("Enter AUTHOR  :", "Search by Book No")
    Adodc2.Recordset.Filter = "AUTHOR ='" & str4 & "'"
End If
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim repp1 As Integer
If Val(Adodc2.Recordset.RecordCount) <= 0 Then
    MsgBox "No more Records to be deleted.", vbInformation, "Confirm"
    Exit Sub
End If
repp1 = MsgBox("You are about to delete 1 record." & vbCrLf & "If you click YES, you won't be able to undo this delete operation." & vbCrLf & "Are you sure you want to delete these record?", vbCritical + vbYesNo, "Confirm Delete")
If repp1 = vbYes Then
    Adodc2.Recordset.Delete
    Adodc2.Recordset.MoveNext
    If Adodc2.Recordset.EOF Then
        Adodc2.Recordset.MoveLast
    End If
    MsgBox "Record has been successfuly deleted.", vbInformation, "Confirm"
End If
End Sub

Private Sub Command9_Click()
Adodc2.Refresh
End Sub

Private Sub DTPicker1_Change()
Adodc4.Recordset.Filter = ""
Adodc4.Recordset.Filter = "DATE_DUE ='" & Format(DTPicker1.Value, "mm/dd/yy") & "'"
End Sub

Private Sub Form_Load()
Me.MousePointer = vbHourglass
DTPicker1.Value = Format(Date, "mm/dd/yy")
Me.Top = 0
Me.Left = 0
Adodc1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc1.RecordSource = "Select * From BARROWERS Order by BARROWERS_NAME"
        Set DataGrid1.DataSource = Adodc1
Adodc2.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc2.RecordSource = "Select * From BOOKS Order by BOOK_TITLE"
        Set DataGrid2.DataSource = Adodc2
Adodc3.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc3.RecordSource = "Select * From BARROWED_BOOKS Where DATE_RETURNED ='NOT YET' Order by BOOK_TITLE"
        Set DataGrid3.DataSource = Adodc3
Adodc4.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc4.RecordSource = "Select * From BARROWED_BOOKS Where DATE_RETURNED ='NOT YET' Order by BOOK_TITLE"
        Set DataGrid4.DataSource = Adodc4
    Adodc4.Recordset.Filter = ""
    Adodc4.Recordset.Filter = "DATE_DUE ='" & Format(DTPicker1.Value, "mm/dd/yy") & "'"
Adodc5.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc5.RecordSource = "Select * From BARROWED_BOOKS Where DATE_RETURNED <>'NOT YET' Order by BOOK_TITLE"
        Set DataGrid5.DataSource = Adodc5
Adodc6.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\DataBase.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=rjbbc")
    Adodc6.RecordSource = "Select * From CATEGORY Order by CATEGORY"
        Set DataGrid6.DataSource = Adodc6
DataGrid1.AllowUpdate = False
DataGrid2.AllowUpdate = False
DataGrid3.AllowUpdate = False
DataGrid4.AllowUpdate = False
DataGrid5.AllowUpdate = False
DataGrid6.AllowUpdate = False
Me.MousePointer = vbDefault
End Sub

