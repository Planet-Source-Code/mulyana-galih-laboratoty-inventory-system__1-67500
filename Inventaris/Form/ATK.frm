VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ATK 
   BackColor       =   &H80000009&
   Caption         =   "Alat Tulis Kantor"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "ATK.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "ATK.frx":74F2
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      Height          =   6015
      Left            =   2400
      TabIndex        =   45
      Top             =   1080
      Width           =   8055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "ATK.frx":79DA
         Height          =   5655
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   9975
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "ID"
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
            DataField       =   "NAMA"
            Caption         =   "NAMA"
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
            DataField       =   "JUMLAH"
            Caption         =   "JUMLAH"
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
            DataField       =   "LOKASI"
            Caption         =   "LOKASI"
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
            DataField       =   "KETERANGAN"
            Caption         =   "KETERANGAN"
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
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2039,811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2099,906
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "TBLATK"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Caption         =   "INFORMASI INVENTARIS ATK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2400
      TabIndex        =   39
      Top             =   1080
      Width           =   8055
      Begin VB.ComboBox ATK4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   360
         ItemData        =   "ATK.frx":79EF
         Left            =   1560
         List            =   "ATK.frx":79F9
         TabIndex        =   4
         Text            =   "Select -------->"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox ATK5 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox ATK3 
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
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox ATK2 
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
         Left            =   1560
         TabIndex        =   33
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox ATK1 
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
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin Project1.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   6120
         TabIndex        =   7
         Top             =   2280
         Width           =   1185
         _extentx        =   2090
         _extenty        =   714
         caption         =   "&Close"
         capalign        =   2
         backstyle       =   2
         gradient        =   3
         cgradient       =   16777215
         cfore           =   11891757
         font            =   "ATK.frx":7A12
         mode            =   0
         value           =   0   'False
         image           =   "ATK.frx":7A3E
         cfhover         =   11891757
         cback           =   16777215
      End
      Begin Project1.lvButtons_H cmdSave 
         Height          =   405
         Left            =   6120
         TabIndex        =   6
         Top             =   1800
         Width           =   1185
         _extentx        =   2090
         _extenty        =   714
         caption         =   "&Save"
         capalign        =   2
         backstyle       =   2
         gradient        =   3
         cgradient       =   16777215
         cfore           =   11891757
         font            =   "ATK.frx":7E78
         mode            =   0
         value           =   0   'False
         image           =   "ATK.frx":7EA4
         cfhover         =   11891757
         cback           =   12632256
         cbhover         =   16777215
         capstyle        =   2
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
         Index           =   6
         Left            =   420
         TabIndex        =   47
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&ID Number:"
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
         Left            =   480
         TabIndex        =   43
         Top             =   465
         Width           =   975
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nama:"
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
         Left            =   900
         TabIndex        =   42
         Top             =   900
         Width           =   555
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Jumlah:"
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
         Left            =   795
         TabIndex        =   41
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Lokasi:"
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
         Left            =   825
         TabIndex        =   40
         Top             =   1800
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   3600
      Left            =   2400
      TabIndex        =   44
      Top             =   1080
      Width           =   8050
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   360
         ItemData        =   "ATK.frx":82FE
         Left            =   1770
         List            =   "ATK.frx":8308
         TabIndex        =   13
         Text            =   "Select -------->"
         Top             =   2400
         Width           =   2535
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
         Left            =   1800
         TabIndex        =   14
         Top             =   2880
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
         Left            =   1770
         TabIndex        =   10
         Top             =   855
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
         Index           =   1
         Left            =   1770
         TabIndex        =   11
         Top             =   1380
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
         Index           =   2
         Left            =   1770
         TabIndex        =   12
         Top             =   1905
         Width           =   3630
      End
      Begin VB.TextBox txtEdBookID 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   3630
      End
      Begin Project1.lvButtons_H cmdUpdate 
         Height          =   405
         Left            =   6120
         TabIndex        =   15
         Top             =   2880
         Width           =   1305
         _extentx        =   2302
         _extenty        =   714
         caption         =   "&Update"
         capalign        =   2
         backstyle       =   2
         gradient        =   3
         cgradient       =   16777215
         cfore           =   11891757
         font            =   "ATK.frx":8321
         mode            =   0
         value           =   0   'False
         image           =   "ATK.frx":834D
         cfhover         =   11891757
         enabled         =   0   'False
         cback           =   16777215
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   405
         Left            =   6120
         TabIndex        =   9
         Top             =   225
         Width           =   1305
         _extentx        =   2302
         _extenty        =   714
         caption         =   "&Search"
         capalign        =   2
         backstyle       =   2
         gradient        =   3
         cgradient       =   16777215
         cfore           =   11891757
         font            =   "ATK.frx":874F
         mode            =   0
         value           =   0   'False
         image           =   "ATK.frx":877B
         imgalign        =   1
         cfhover         =   11891757
         cback           =   16777215
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&ID Number:"
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
         Left            =   600
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Lokasi:"
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
         Left            =   1005
         TabIndex        =   52
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Jumlah:"
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
         Left            =   975
         TabIndex        =   51
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nama:"
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
         Left            =   1080
         TabIndex        =   50
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&ID Number:"
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
         Left            =   660
         TabIndex        =   49
         Top             =   960
         Width           =   975
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
         Index           =   4
         Left            =   600
         TabIndex        =   48
         Top             =   3000
         Width           =   1050
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   4455
      Left            =   2400
      TabIndex        =   21
      Top             =   1080
      Width           =   8055
      Begin VB.Frame fmeDel 
         BackColor       =   &H8000000E&
         Caption         =   "ATK Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2520
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   7575
         Begin VB.PictureBox picContainer 
            Appearance      =   0  'Flat
            BackColor       =   &H00875B25&
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   1590
            ScaleHeight     =   1395
            ScaleWidth      =   5595
            TabIndex        =   25
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
               ForeColor       =   &H0000FF00&
               Height          =   210
               Index           =   4
               Left            =   75
               TabIndex        =   30
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
               ForeColor       =   &H0000FF00&
               Height          =   210
               Index           =   3
               Left            =   75
               TabIndex        =   29
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
               ForeColor       =   &H0000FF00&
               Height          =   210
               Index           =   2
               Left            =   75
               TabIndex        =   28
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
               ForeColor       =   &H0000FF00&
               Height          =   210
               Index           =   1
               Left            =   75
               TabIndex        =   27
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
               ForeColor       =   &H0000FF00&
               Height          =   210
               Index           =   0
               Left            =   75
               TabIndex        =   26
               Top             =   45
               Width           =   120
            End
         End
         Begin Project1.lvButtons_H cmdDel 
            Height          =   405
            Left            =   5880
            TabIndex        =   18
            Top             =   1920
            Width           =   1305
            _extentx        =   2302
            _extenty        =   714
            caption         =   "&Delete"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   16777215
            cfore           =   11891757
            font            =   "ATK.frx":9455
            mode            =   0
            value           =   0   'False
            image           =   "ATK.frx":9481
            imgalign        =   1
            cfhover         =   11891757
            enabled         =   0   'False
            cback           =   16777215
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
            Index           =   11
            Left            =   450
            TabIndex        =   38
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Lokasi:"
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
            Left            =   870
            TabIndex        =   37
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Jumlah:"
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
            Left            =   855
            TabIndex        =   36
            Top             =   960
            Width           =   660
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Nama:"
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
            Left            =   945
            TabIndex        =   35
            Top             =   720
            Width           =   555
         End
         Begin VB.Label lblborInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&ID Number:"
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
            Left            =   525
            TabIndex        =   34
            Top             =   480
            Width           =   975
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
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   4695
            Width           =   1050
         End
      End
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
         TabIndex        =   22
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
            TabIndex        =   16
            Top             =   240
            Width           =   3915
         End
         Begin Project1.lvButtons_H cmdSearchDel 
            Height          =   405
            Left            =   5880
            TabIndex        =   17
            Top             =   240
            Width           =   1305
            _extentx        =   2302
            _extenty        =   714
            caption         =   "&Search"
            capalign        =   2
            backstyle       =   2
            gradient        =   3
            cgradient       =   16777215
            cfore           =   11891757
            font            =   "ATK.frx":A15B
            mode            =   0
            value           =   0   'False
            image           =   "ATK.frx":A187
            imgalign        =   1
            cfhover         =   11891757
            cback           =   16777215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID Number:"
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
            Left            =   435
            TabIndex        =   23
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan data yang diminta dibawah ini, sesuai dengan spesifikasi alat tulis kantor (ATK) tersebut."
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
      TabIndex        =   20
      Top             =   360
      Width           =   9225
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   360
      Picture         =   "ATK.frx":AE61
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input Data ATK"
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
      MouseIcon       =   "ATK.frx":BB2B
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Data ATK"
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
      MouseIcon       =   "ATK.frx":BE35
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Data ATK"
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
      MouseIcon       =   "ATK.frx":C13F
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1920
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C7BDAD&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   1155
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "ATK.frx":C449
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "ATK"
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
Adodc1.Refresh
Exit Sub
a:
MsgBox "Tidak ada data yang dihapus", vbCritical + vbOKOnly, "Pesan penghapusan"
End Sub

Private Sub cmdSave_Click()

Dim Mcari As String
Dim X As Integer
Mcari = "ID= '" & ATK1.Text & "'"

If ATK1.Text = "" Or ATK2.Text = "" Then
MsgBox "Masih ada data yang belum di Isi", vbOKOnly + vbCritical, "Pesan Pengisian"
End If
On Error Resume Next
With Adodc1.Recordset
.Find Mcari
If Not .EOF Then
X = MsgBox("Maaf untuk nomor ID [" & ATK1.Text & "] sudah dimasukan", vbOKOnly + vbInformation, ":. Pesan")
Exit Sub
End If
End With
With Adodc1.Recordset
.AddNew
!Id = UCase(ATK1.Text)
!NAMA = UCase(ATK2.Text)
!JUMLAH = UCase(ATK3.Text)
!LOKASI = UCase(ATK4.Text)
!KETERANGAN = UCase(ATK5.Text)
.Update


On Error GoTo 0
ATK1.Text = ""
ATK2.Text = ""
ATK3.Text = ""
ATK4.Text = ""
ATK5.Text = ""
End With

End Sub


Private Sub cmdSearch_Click()
On Error GoTo NotFound
        If Trim(txtEdBookID.Text) = "" Then
            Exit Sub
        End If
        
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "No ID yang dicari tidak ditemukan", vbOKOnly + vbExclamation, "Inventory System"
            Exit Sub
        End If
    
    Adodc1.Refresh
    Adodc1.Recordset.Find ("ID = '" & Trim(txtEdBookID.Text) & "'")

   Text(0).Text = Adodc1.Recordset.Fields("ID")
   Text(1).Text = Adodc1.Recordset.Fields("NAMA")
   Text(2).Text = Adodc1.Recordset.Fields("JUMLAH")
   Combo2.Text = Adodc1.Recordset.Fields("LOKASI")
   Text(3).Text = Adodc1.Recordset.Fields("KETERANGAN")

   Text(0).Enabled = True
   Text(1).Enabled = True
   Text(2).Enabled = True
   Combo2.Enabled = True
   Text(3).Enabled = True
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

   lblInfo(0).Caption = Adodc1.Recordset.Fields("ID")
   lblInfo(1).Caption = Adodc1.Recordset.Fields("NAMA")
   lblInfo(2).Caption = Adodc1.Recordset.Fields("JUMLAH")
   lblInfo(3).Caption = Adodc1.Recordset.Fields("LOKASI")
   lblInfo(4).Caption = Adodc1.Recordset.Fields("KETERANGAN")
   
   cmdDel.Enabled = True
   
  

Exit Sub
NotFound:
MsgBox "Data yang dicari tidak ditemukan", vbOKCancel + vbInformation, "Inventory System"
End Sub

Private Sub cmdUpdate_Click()
With Adodc1.Recordset
!Id = UCase(Text(0).Text)
!NAMA = UCase(Text(1).Text)
!JUMLAH = UCase(Text(2).Text)
!LOKASI = UCase(Combo2.Text)
!KETERANGAN = UCase(Text(3).Text)
.Update
End With
cmdUpdate.Enabled = False
Frame2_Kosong
End Sub

Private Sub Form_Activate()
Unload INPUTKOMPUTER
Unload FASILITAS
Unload PERALATAN

End Sub

Private Sub Label1_Click(Index As Integer)

Select Case Index

Case 0
Frame3.Visible = True
ATK1.SetFocus
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
ATK.Refresh
Frame1_Kosong

Case 2
Frame3.Visible = False
Frame1.Visible = True
Frame2.Visible = False
Frame5.Visible = False
Adodc1.Refresh
ATK.Refresh
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
txtEdBookID.Text = ""
Text(0).Text = ""
Text(1).Text = ""
Text(2).Text = ""
Combo2.Text = "Select -------->"
Text(3).Text = ""
Text(0).Enabled = False
Text(1).Enabled = False
Text(2).Enabled = False
Text(3).Enabled = False
Combo2.Enabled = False
End Sub

Public Sub Frame1_Kosong()
txtIdDel.Text = ""
lblInfo(0).Caption = "--"
lblInfo(1).Caption = "--"
lblInfo(2).Caption = "--"
lblInfo(3).Caption = "--"
lblInfo(4).Caption = "--"
cmdDel.Enabled = False
End Sub

