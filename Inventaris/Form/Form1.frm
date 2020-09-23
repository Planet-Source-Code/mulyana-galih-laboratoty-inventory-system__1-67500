VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ADMINS 
   BackColor       =   &H80000009&
   Caption         =   "Administrator System"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database\USER LOGIN.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database\USER LOGIN.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TBL_LOGIN"
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
      Height          =   3975
      Left            =   2640
      TabIndex        =   44
      Top             =   1080
      Width           =   7815
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form1.frx":74F2
         Height          =   3615
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         ColumnHeaders   =   -1  'True
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   18
         TabAction       =   1
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Username"
            Caption         =   "Username"
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
            DataField       =   "Name"
            Caption         =   "Name"
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
            DataField       =   "Password"
            Caption         =   "Password"
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
            DataField       =   "Alamat"
            Caption         =   "Alamat"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   3975
      Left            =   2640
      TabIndex        =   15
      Top             =   1080
      Width           =   7815
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
         TabIndex        =   16
         Top             =   480
         Width           =   7335
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
            TabIndex        =   17
            Top             =   240
            Width           =   4035
         End
         Begin Project1.lvButtons_H cmdSearchDel 
            Height          =   405
            Left            =   5880
            TabIndex        =   18
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
            Image           =   "Form1.frx":7507
            cBack           =   16777215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
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
            Left            =   495
            TabIndex        =   19
            Top             =   360
            Width           =   915
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
         Height          =   2040
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   7335
         Begin VB.PictureBox picContainer 
            Appearance      =   0  'Flat
            BackColor       =   &H00875B25&
            ForeColor       =   &H80000008&
            Height          =   945
            Left            =   1560
            ScaleHeight     =   915
            ScaleWidth      =   5595
            TabIndex        =   21
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
               Index           =   2
               Left            =   75
               TabIndex        =   24
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
               Index           =   1
               Left            =   75
               TabIndex        =   23
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
               Index           =   0
               Left            =   75
               TabIndex        =   22
               Top             =   45
               Width           =   120
            End
         End
         Begin Project1.lvButtons_H cmdDel 
            Height          =   405
            Left            =   5880
            TabIndex        =   25
            Top             =   1440
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
            Image           =   "Form1.frx":81E1
            Enabled         =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Alamat:"
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
            Left            =   870
            TabIndex        =   30
            Top             =   960
            Width           =   645
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Nama Lengkap:"
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
            Left            =   150
            TabIndex        =   29
            Top             =   720
            Width           =   1350
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Username:"
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
            Left            =   585
            TabIndex        =   28
            Top             =   480
            Width           =   915
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
            TabIndex        =   27
            Top             =   4215
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
            TabIndex        =   26
            Top             =   4695
            Width           =   1050
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   3360
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   7815
      Begin VB.TextBox ADMIN 
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
         Height          =   780
         Index           =   3
         Left            =   1890
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2400
         Width           =   3630
      End
      Begin VB.TextBox ADMIN 
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
         Left            =   1890
         TabIndex        =   9
         Top             =   855
         Width           =   3630
      End
      Begin VB.TextBox ADMIN 
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
         Index           =   1
         Left            =   1890
         TabIndex        =   10
         Top             =   1380
         Width           =   3630
      End
      Begin VB.TextBox ADMIN 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1890
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1905
         Width           =   3630
      End
      Begin VB.TextBox AdminID 
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
         Left            =   1890
         TabIndex        =   7
         Top             =   240
         Width           =   3630
      End
      Begin Project1.lvButtons_H cmdUpdate 
         Height          =   405
         Left            =   6120
         TabIndex        =   13
         Top             =   2760
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
         Image           =   "Form1.frx":8EBB
         Enabled         =   0   'False
         cBack           =   16777215
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   405
         Left            =   6120
         TabIndex        =   8
         Top             =   225
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
         Image           =   "Form1.frx":92BB
         cBack           =   16777215
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Alamat:"
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
         Left            =   1200
         TabIndex        =   43
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
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
         Left            =   945
         TabIndex        =   42
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nama Lengkap:"
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
         Left            =   480
         TabIndex        =   41
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Username:"
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
         Left            =   915
         TabIndex        =   40
         Top             =   960
         Width           =   915
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   8160
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   8160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
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
         Left            =   900
         TabIndex        =   14
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Caption         =   "INFORMASI ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2640
      TabIndex        =   31
      Top             =   1080
      Width           =   7815
      Begin VB.TextBox Text 
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
         Height          =   810
         Index           =   3
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox Text 
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
      Begin VB.TextBox Text 
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
      Begin Project1.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   6000
         TabIndex        =   6
         Top             =   2160
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
         cBhover         =   -2147483639
         cGradient       =   -2147483639
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Form1.frx":9F95
         cBack           =   16777215
      End
      Begin Project1.lvButtons_H cmdSave 
         Height          =   405
         Left            =   6000
         TabIndex        =   5
         Top             =   1560
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
         Image           =   "Form1.frx":A3CE
         cBack           =   12632256
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Username:"
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
         Left            =   540
         TabIndex        =   35
         Top             =   465
         Width           =   915
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nama Lengkap:"
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
         Left            =   105
         TabIndex        =   34
         Top             =   900
         Width           =   1350
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
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
         Left            =   570
         TabIndex        =   33
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Alamat:"
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
         Left            =   810
         TabIndex        =   32
         Top             =   1800
         Width           =   645
      End
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan data yang diminta dibawah ini, sesuai dengan spesifikasi pada form tersebut."
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
      TabIndex        =   39
      Top             =   360
      Width           =   9225
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   360
      Picture         =   "Form1.frx":A826
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input Data Admin"
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
      MouseIcon       =   "Form1.frx":B4F0
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   1200
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Data Admin"
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
      MouseIcon       =   "Form1.frx":B7FA
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Data Admin"
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
      MouseIcon       =   "Form1.frx":BB04
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   1920
      Width           =   1755
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "Form1.frx":BE0E
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
      Left            =   360
      Top             =   1155
      Visible         =   0   'False
      Width           =   2085
   End
End
Attribute VB_Name = "ADMINS"
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
Adodc1.Refresh
Exit Sub
a:
MsgBox "Tidak ada data yang dihapus", vbCritical + vbOKOnly, "Pesan penghapusan"
End Sub

Private Sub cmdSave_Click()

Dim Mcari As String
Dim X As Integer
Mcari = "Username= '" & Text(0).Text & "'"

If Text(0).Text = "" Or Text(1).Text = "" Then
MsgBox "Masih ada data yang belum di Isi", vbOKOnly + vbCritical, "Pesan Pengisian"
End If
On Error Resume Next
With Adodc1.Recordset
.Find Mcari
If Not .EOF Then
X = MsgBox("Maaf untuk username [" & Text(0).Text & "] sudah dimasukan", vbOKOnly + vbInformation, ":. Pesan")
Exit Sub
End If
End With
With Adodc1.Recordset
.AddNew
!UserName = Text(0).Text
!Name = UCase(Text(1).Text)
!Password = Text(2).Text
!Alamat = UCase(Text(3).Text)
.Update


On Error GoTo 0
Text(0).Text = ""
Text(1).Text = ""
Text(2).Text = ""
Text(3).Text = ""

End With
End Sub

Private Sub cmdSearch_Click()
On Error GoTo NotFound
        If Trim(AdminID.Text) = "" Then
            Exit Sub
        End If
        
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Username yang dicari tidak ditemukan", vbOKOnly + vbExclamation, "Inventory System"
            Exit Sub
        End If
    
    Adodc1.Refresh
    Adodc1.Recordset.Find ("Username = '" & Trim(AdminID.Text) & "'")

   ADMIN(0).Text = Adodc1.Recordset.Fields("Username")
   ADMIN(1).Text = Adodc1.Recordset.Fields("Name")
   ADMIN(2).Text = Adodc1.Recordset.Fields("Password")
   ADMIN(3).Text = Adodc1.Recordset.Fields("Alamat")

   ADMIN(0).Enabled = True
   ADMIN(1).Enabled = True
   ADMIN(2).Enabled = True
   ADMIN(3).Enabled = True
   cmdUpdate.Enabled = True


Exit Sub
NotFound:
MsgBox "Username yang dicari tidak ditemukan", vbOKCancel + vbInformation, "Inventory System"
End Sub

Private Sub cmdSearchDel_Click()
On Error GoTo NotFound
        If Trim(txtIdDel.Text) = "" Then
            Exit Sub
        End If
        
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Username yang dicari tidak ditemukan", vbOKOnly + vbExclamation, "Inventory System"
            Exit Sub
        End If
    
    Adodc1.Refresh
    Adodc1.Recordset.Find ("Username = '" & Trim(txtIdDel.Text) & "'")

   lblInfo(0).Caption = Adodc1.Recordset.Fields("Username")
   lblInfo(1).Caption = Adodc1.Recordset.Fields("Name")
   lblInfo(2).Caption = Adodc1.Recordset.Fields("Alamat")
   cmdDel.Enabled = True
    
Exit Sub
NotFound:
MsgBox "Username yang dicari tidak ditemukan", vbOKCancel + vbInformation, "Inventory System"
End Sub

Private Sub cmdUpdate_Click()
With Adodc1.Recordset
!UserName = ADMIN(0).Text
!Name = UCase(ADMIN(1).Text)
!Password = ADMIN(2).Text
!Alamat = UCase(ADMIN(3).Text)
.Update
End With
cmdUpdate.Enabled = False
Frame2_Kosong
End Sub

Private Sub Label1_Click(Index As Integer)

Select Case Index

Case 0
Frame3.Visible = True
Text(0).SetFocus
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
Frame1_Kosong
ADMINS.Refresh

Case 2
Frame3.Visible = False
Frame1.Visible = True
Frame2.Visible = False
Frame5.Visible = False
Adodc1.Refresh
ADMINS.Refresh
Frame2_Kosong
End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveShape(Shape1, Label1(Index))
End Sub
Public Function moveShape(shape As Object, Cntrl As Object)
        shape.Visible = True
        shape.Move Cntrl.Left - 150, Cntrl.Top - 60, 2000, 300
End Function

Public Sub Frame2_Kosong()
AdminID.Text = ""
ADMIN(0).Text = ""
ADMIN(1).Text = ""
ADMIN(2).Text = ""
ADMIN(3).Text = ""
ADMIN(0).Enabled = False
ADMIN(1).Enabled = False
ADMIN(2).Enabled = False
ADMIN(3).Enabled = False
End Sub

Public Sub Frame1_Kosong()
txtIdDel.Text = ""
lblInfo(0).Caption = "--"
lblInfo(1).Caption = "--"
lblInfo(2).Caption = "--"
cmdDel.Enabled = False
End Sub

