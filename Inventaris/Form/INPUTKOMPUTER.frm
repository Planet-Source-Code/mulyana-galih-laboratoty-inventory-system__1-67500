VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form INPUTKOMPUTER 
   BackColor       =   &H80000009&
   Caption         =   "Input Komputer"
   ClientHeight    =   7710
   ClientLeft      =   1530
   ClientTop       =   1905
   ClientWidth     =   8445
   Icon            =   "INPUTKOMPUTER.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2280
      Top             =   7440
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   7920
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database\INVENTDB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database\INVENTDB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TBLKOMP"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      Height          =   6015
      Left            =   2280
      TabIndex        =   74
      Top             =   1080
      Width           =   8055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "INPUTKOMPUTER.frx":74F2
         Height          =   5655
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   -2147483647
         HeadLines       =   1
         RowHeight       =   18
         TabAction       =   1
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "NUMBER"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "mb"
            Caption         =   "MOTHERBOARD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "pc"
            Caption         =   "PROCESSOR"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "hdd"
            Caption         =   "HARDDISK"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "ram"
            Caption         =   "RAM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "vga"
            Caption         =   "VGA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "sc"
            Caption         =   "SOUNDCARD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "cd"
            Caption         =   "CDROM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "lc"
            Caption         =   "LAN CARD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "mon"
            Caption         =   "MONITOR"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "ket"
            Caption         =   "KET"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   6015
      Left            =   2280
      TabIndex        =   47
      Top             =   1080
      Width           =   8055
      Begin VB.Frame Frame4 
         BackColor       =   &H80000009&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   56
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtIdDel 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00B5742D&
            Height          =   420
            Left            =   1560
            TabIndex        =   29
            Top             =   240
            Width           =   3915
         End
         Begin Project1.lvButtons_H cmdSearchDel 
            Height          =   405
            Left            =   5880
            TabIndex        =   31
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   714
            Caption         =   "&Search"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   11891757
            cFHover         =   11891757
            cGradient       =   16777215
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            ImgAlign        =   1
            Image           =   "INPUTKOMPUTER.frx":7507
            cBack           =   16777215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID Komputer:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   285
            TabIndex        =   57
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.Frame fmeDel 
         BackColor       =   &H8000000E&
         Caption         =   "Computer Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   240
         TabIndex        =   48
         Top             =   1680
         Width           =   7575
         Begin VB.PictureBox picContainer 
            Appearance      =   0  'Flat
            BackColor       =   &H00875B25&
            ForeColor       =   &H80000008&
            Height          =   2865
            Left            =   1590
            ScaleHeight     =   2835
            ScaleWidth      =   5595
            TabIndex        =   49
            Top             =   360
            Width           =   5625
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   10
               Left            =   75
               TabIndex        =   87
               Top             =   2520
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   9
               Left            =   75
               TabIndex        =   73
               Top             =   2280
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   8
               Left            =   75
               TabIndex        =   72
               Top             =   2040
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   7
               Left            =   75
               TabIndex        =   69
               Top             =   1800
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   6
               Left            =   75
               TabIndex        =   68
               Top             =   1560
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   0
               Left            =   75
               TabIndex        =   55
               Top             =   45
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   1
               Left            =   75
               TabIndex        =   54
               Top             =   315
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   2
               Left            =   75
               TabIndex        =   53
               Top             =   570
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   3
               Left            =   75
               TabIndex        =   52
               Top             =   840
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   4
               Left            =   75
               TabIndex        =   51
               Top             =   1095
               Width           =   120
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   5
               Left            =   75
               TabIndex        =   50
               Top             =   1365
               Width           =   120
            End
         End
         Begin Project1.lvButtons_H cmdDel 
            Height          =   405
            Left            =   5880
            TabIndex        =   30
            Top             =   3360
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   714
            Caption         =   "&Delete"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   11891757
            cFHover         =   11891757
            cGradient       =   16777215
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            ImgAlign        =   1
            Image           =   "INPUTKOMPUTER.frx":81E1
            Enabled         =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Keterangan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   480
            TabIndex        =   86
            Top             =   2880
            Width           =   1050
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&RAM:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   1040
            TabIndex        =   71
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "M&onitor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   840
            TabIndex        =   70
            Top             =   2640
            Width           =   705
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Keterangan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   300
            TabIndex        =   67
            Top             =   4695
            Width           =   1050
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "M&onitor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   690
            TabIndex        =   66
            Top             =   4215
            Width           =   705
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Lan Card:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   720
            TabIndex        =   65
            Top             =   2400
            Width           =   840
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Sound Card:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   480
            TabIndex        =   64
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&ID Komputer:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   375
            TabIndex        =   63
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Motherboard:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   360
            TabIndex        =   62
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Processor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   600
            TabIndex        =   61
            Top             =   960
            Width           =   915
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Harddisk:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   675
            TabIndex        =   60
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&VGA:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   1050
            TabIndex        =   59
            Top             =   1680
            Width           =   450
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&CDROM:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   765
            TabIndex        =   58
            Top             =   2160
            Width           =   750
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   6240
      Left            =   2280
      TabIndex        =   43
      Top             =   1080
      Width           =   8050
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   4
         Left            =   1410
         TabIndex        =   21
         Top             =   2720
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   10
         Left            =   1410
         TabIndex        =   27
         Top             =   5640
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   7
         Left            =   1410
         TabIndex        =   24
         Top             =   4150
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   8
         Left            =   1410
         TabIndex        =   25
         Top             =   4645
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   9
         Left            =   1410
         TabIndex        =   26
         Top             =   5130
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   3
         Left            =   1410
         TabIndex        =   20
         Top             =   2240
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   5
         Left            =   1410
         TabIndex        =   22
         Top             =   3200
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   6
         Left            =   1410
         TabIndex        =   23
         Top             =   3680
         Width           =   3630
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Left            =   1410
         TabIndex        =   15
         Top             =   240
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   2
         Left            =   1410
         TabIndex        =   19
         Top             =   1755
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   420
         Index           =   1
         Left            =   1410
         TabIndex        =   18
         Top             =   1280
         Width           =   3630
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   0
         Left            =   1410
         TabIndex        =   17
         Top             =   840
         Width           =   3630
      End
      Begin Project1.lvButtons_H cmdUpdate 
         Height          =   405
         Left            =   5400
         TabIndex        =   28
         Top             =   5640
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   714
         Caption         =   "&Update"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   11891757
         cFHover         =   11891757
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "INPUTKOMPUTER.frx":8EBB
         Enabled         =   0   'False
         cBack           =   16777215
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   405
         Left            =   5400
         TabIndex        =   16
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   714
         Caption         =   "&Search"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   11891757
         cFHover         =   11891757
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "INPUTKOMPUTER.frx":92BB
         cBack           =   16777215
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&RAM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   760
         TabIndex        =   89
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Keterangan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   180
         TabIndex        =   85
         Top             =   5640
         Width           =   1050
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M&onitor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   570
         TabIndex        =   84
         Top             =   5160
         Width           =   705
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Lan Card:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   435
         TabIndex        =   83
         Top             =   4690
         Width           =   840
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sound Card:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   210
         TabIndex        =   82
         Top             =   3840
         Width           =   1065
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&ID Komputer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   135
         TabIndex        =   81
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Motherboard:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   120
         TabIndex        =   80
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Processor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   345
         TabIndex        =   79
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Harddisk:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   435
         TabIndex        =   78
         Top             =   2400
         Width           =   825
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&VGA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   810
         TabIndex        =   77
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CDROM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   525
         TabIndex        =   76
         Top             =   4320
         Width           =   750
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Komputer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   210
         TabIndex        =   44
         Top             =   240
         Width           =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   8160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   8160
         Y1              =   735
         Y2              =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Caption         =   "INFORMASI INVENTARIS KOMPUTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   2280
      TabIndex        =   32
      Top             =   1080
      Width           =   8055
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   5
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   5
         Left            =   1560
         TabIndex        =   6
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   6
         Left            =   1560
         TabIndex        =   7
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   7
         Left            =   1560
         TabIndex        =   8
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   8
         Left            =   1560
         TabIndex        =   9
         Top             =   4200
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Index           =   9
         Left            =   1560
         TabIndex        =   10
         Top             =   4680
         Width           =   3735
      End
      Begin VB.TextBox kompi 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   735
         Index           =   10
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   5160
         Width           =   3735
      End
      Begin Project1.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   6120
         TabIndex        =   13
         Top             =   5400
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   714
         Caption         =   "&Close"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   11891757
         cFHover         =   11891757
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "INPUTKOMPUTER.frx":9F95
         cBack           =   16777215
      End
      Begin Project1.lvButtons_H cmdSave 
         Height          =   405
         Left            =   6120
         TabIndex        =   12
         Top             =   4920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   714
         Caption         =   "&Save"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   11891757
         cFHover         =   11891757
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   3
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         Image           =   "INPUTKOMPUTER.frx":A3CE
         cBack           =   12632256
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&RAM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   960
         TabIndex        =   88
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CDROM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   42
         Top             =   3720
         Width           =   750
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&VGA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1005
         TabIndex        =   41
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Harddisk:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   40
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Processor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   39
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Motherboard:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   38
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&ID Komputer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   37
         Top             =   465
         Width           =   1125
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sound Card:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   405
         TabIndex        =   36
         Top             =   3240
         Width           =   1065
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Lan Card:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   630
         TabIndex        =   35
         Top             =   4200
         Width           =   840
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M&onitor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   765
         TabIndex        =   34
         Top             =   4800
         Width           =   705
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Keterangan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   375
         TabIndex        =   33
         Top             =   5280
         Width           =   1050
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   2
      Left            =   480
      MouseIcon       =   "INPUTKOMPUTER.frx":A826
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   1
      Left            =   480
      MouseIcon       =   "INPUTKOMPUTER.frx":AB30
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input Data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   0
      Left            =   480
      MouseIcon       =   "INPUTKOMPUTER.frx":AE3A
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   360
      Picture         =   "INPUTKOMPUTER.frx":B144
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan data yang diminta dibawah ini, sesuai dengan spesifikasi komputer tersebut."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   9225
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "INPUTKOMPUTER.frx":BE0E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C7BDAD&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   1150
      Visible         =   0   'False
      Width           =   1845
   End
End
Attribute VB_Name = "INPUTKOMPUTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDel_Click()
On Error GoTo a
With Adodc1.Recordset
.Delete
End With
txtIdDel.Text = ""
lblInfo(0).Caption = "--"
lblInfo(1).Caption = "--"
lblInfo(2).Caption = "--"
lblInfo(3).Caption = "--"
lblInfo(4).Caption = "--"
lblInfo(5).Caption = "--"
lblInfo(6).Caption = "--"
lblInfo(7).Caption = "--"
lblInfo(8).Caption = "--"
lblInfo(9).Caption = "--"
lblInfo(10).Caption = "--"
Adodc1.Refresh
Exit Sub
a:
MsgBox "Tidak ada data yang dihapus", vbCritical + vbOKOnly, "Pesan penghapusan"
End Sub

Private Sub cmdSave_Click()
Dim Mcari As String
Dim X As Integer
Mcari = "id= '" & kompi(0).Text & "'"

If kompi(0).Text = "" Or kompi(1).Text = "" Then
MsgBox "Masih ada data yang belum di Isi", vbOKOnly + vbCritical, "Pesan Pengisian"
End If
On Error Resume Next
With Adodc1.Recordset
.Find Mcari
If Not .EOF Then
X = MsgBox("Maaf untuk nomor ID [" & kompi(0).Text & "] sudah dimasukan", vbOKOnly + vbInformation, ":. Pesan")
Exit Sub
End If
End With
With Adodc1.Recordset
.AddNew
!Id = UCase(kompi(0).Text)
!mb = UCase(kompi(1).Text)
!pc = UCase(kompi(2).Text)
!hdd = UCase(kompi(3).Text)
!ram = UCase(kompi(4).Text)
!vga = UCase(kompi(5).Text)
!sc = UCase(kompi(6).Text)
!cd = UCase(kompi(7).Text)
!lc = UCase(kompi(8).Text)
!mon = UCase(kompi(9).Text)
!ket = UCase(kompi(10).Text)
.Update

On Error GoTo 0
kompi(0).Text = ""
kompi(1).Text = ""
kompi(2).Text = ""
kompi(3).Text = ""
kompi(4).Text = ""
kompi(5).Text = ""
kompi(6).Text = ""
kompi(7).Text = ""
kompi(8).Text = ""
kompi(9).Text = ""
kompi(10).Text = ""

End With
End Sub


Private Sub cmdSearch_Click()
On Error GoTo NotFound
        If Trim(txtID.Text) = "" Then
            Exit Sub
        End If
        
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "No ID yang dicari tidak ditemukan", vbOKOnly + vbExclamation, "Inventory System"
            Exit Sub
        End If
    
    Adodc1.Refresh
    Adodc1.Recordset.Find ("ID = '" & Trim(txtID.Text) & "'")

   Text(0).Text = Adodc1.Recordset.Fields("id")
   Text(1).Text = Adodc1.Recordset.Fields("mb")
   Text(2).Text = Adodc1.Recordset.Fields("pc")
   Text(3).Text = Adodc1.Recordset.Fields("hdd")
   Text(4).Text = Adodc1.Recordset.Fields("ram")
   Text(5).Text = Adodc1.Recordset.Fields("vga")
   Text(6).Text = Adodc1.Recordset.Fields("sc")
   Text(7).Text = Adodc1.Recordset.Fields("cd")
   Text(8).Text = Adodc1.Recordset.Fields("lc")
   Text(9).Text = Adodc1.Recordset.Fields("mon")
   Text(10).Text = Adodc1.Recordset.Fields("ket")
  
   Text(0).Enabled = True
   Text(1).Enabled = True
   Text(2).Enabled = True
   Text(3).Enabled = True
   Text(4).Enabled = True
   Text(5).Enabled = True
   Text(6).Enabled = True
   Text(7).Enabled = True
   Text(8).Enabled = True
   Text(9).Enabled = True
   Text(10).Enabled = True
      
   cmdUpdate.Enabled = True


Exit Sub
NotFound:
MsgBox "Data yang dicari tidak ditemukan", vbOKCancel + vbInformation, "Inventory System"

End Sub

Private Sub cmdSearchDel_Click()
On Error GoTo NotFound
        If Trim(txtIdDel.Text) = "" Then
            Exit Sub
        End If
        
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "No ID yang dicari tidak ditemukan", vbOKOnly + vbExclamation, "Inventory System"
            Exit Sub
        End If
    
    Adodc1.Refresh
    Adodc1.Recordset.Find ("ID = '" & Trim(txtIdDel.Text) & "'")

   lblInfo(0).Caption = Adodc1.Recordset.Fields("id")
   lblInfo(1).Caption = Adodc1.Recordset.Fields("mb")
   lblInfo(2).Caption = Adodc1.Recordset.Fields("pc")
   lblInfo(3).Caption = Adodc1.Recordset.Fields("hdd")
   lblInfo(4).Caption = Adodc1.Recordset.Fields("ram")
   lblInfo(5).Caption = Adodc1.Recordset.Fields("vga")
   lblInfo(6).Caption = Adodc1.Recordset.Fields("sc")
   lblInfo(7).Caption = Adodc1.Recordset.Fields("cd")
   lblInfo(8).Caption = Adodc1.Recordset.Fields("lc")
   lblInfo(9).Caption = Adodc1.Recordset.Fields("mon")
   lblInfo(10).Caption = Adodc1.Recordset.Fields("ket")
  
   cmdDel.Enabled = True
   
  

Exit Sub
NotFound:
MsgBox "Data yang dicari tidak ditemukan", vbOKCancel + vbInformation, "Inventory System"
End Sub

Private Sub cmdUpdate_Click()
With Adodc1.Recordset

!Id = UCase(Text(0).Text)
!mb = UCase(Text(1).Text)
!pc = UCase(Text(2).Text)
!hdd = UCase(Text(3).Text)
!ram = UCase(Text(4).Text)
!vga = UCase(Text(5).Text)
!sc = UCase(Text(6).Text)
!cd = UCase(Text(7).Text)
!lc = UCase(Text(8).Text)
!mon = UCase(Text(9).Text)
!ket = UCase(Text(10).Text)

.Update
End With
cmdUpdate.Enabled = False
Frame2_Kosong
End Sub

Private Sub Form_Activate()
Unload ATK
Unload FASILITAS
Unload PERALATAN
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index

Case 0
Frame3.Visible = True
kompi(0).SetFocus
Frame1.Visible = False
Frame2.Visible = False
Frame5.Visible = False
Adodc1.Refresh
Frame1.Refresh
Frame2_Kosong
Frame1_Kosong

Case 1
Frame3.Visible = False
Frame1.Visible = False
Frame2.Visible = True
Frame5.Visible = False
Adodc1.Refresh
INPUTKOMPUTER.Refresh
Frame1_Kosong

Case 2
Frame3.Visible = False
Frame1.Visible = True
Frame2.Visible = False
Frame5.Visible = False
Adodc1.Refresh
INPUTKOMPUTER.Refresh
Frame2_Kosong
End Select

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveShape(Shape1, Label1(Index))
End Sub
Public Function moveShape(shape As Object, Cntrl As Object)
        shape.Visible = True
        shape.Move Cntrl.Left - 150, Cntrl.Top - 60, 1845, 300
End Function

Public Sub Frame2_Kosong()
txtID.Text = ""
Text(0).Text = ""
Text(1).Text = ""
Text(2).Text = ""
Text(3).Text = ""
Text(4).Text = ""
Text(5).Text = ""
Text(6).Text = ""
Text(7).Text = ""
Text(8).Text = ""
Text(9).Text = ""
Text(10).Text = ""

Text(0).Enabled = False
Text(1).Enabled = False
Text(2).Enabled = False
Text(3).Enabled = False
Text(4).Enabled = False
Text(5).Enabled = False
Text(6).Enabled = False
Text(7).Enabled = False
Text(8).Enabled = False
Text(9).Enabled = False
Text(10).Enabled = False

End Sub

Public Sub Frame1_Kosong()
txtIdDel.Text = ""
lblInfo(0).Caption = "--"
lblInfo(1).Caption = "--"
lblInfo(2).Caption = "--"
lblInfo(3).Caption = "--"
lblInfo(4).Caption = "--"
lblInfo(5).Caption = "--"
lblInfo(6).Caption = "--"
lblInfo(7).Caption = "--"
lblInfo(8).Caption = "--"
lblInfo(9).Caption = "--"
lblInfo(10).Caption = "--"
cmdDel.Enabled = False
End Sub

