VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form LOGIN 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   4650
   ClientTop       =   3645
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Roman"
      Size            =   8.25
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000C&
   Icon            =   "LOGIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LOGIN.frx":0E42
   ScaleHeight     =   2145
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   600
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   1680
   End
   Begin Project1.lvButtons_H SignUp 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "SignUp"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "LOGIN.frx":22FB
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      Caption         =   "Daftar Admin"
      ColumnCount     =   2
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
         DataField       =   "Password"
         Caption         =   "Password"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   """Rp""#.##0"
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
      EndProperty
   End
   Begin Project1.lvButtons_H btnlogin 
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Login"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "LOGIN.frx":2310
      cBack           =   -2147483633
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin Project1.lvButtons_H btnexit 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "E&xit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "LOGIN.frx":27C9
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H Delete 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Delete"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CLOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   130
      TabIndex        =   10
      Top             =   1720
      Width           =   1095
   End
   Begin VB.Label marq1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ">> INVENTORY SYSTEM <<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public LoginSukses As Boolean

Private orgcaption As String   'stores the original caption
Private captionlen As Integer  'stores the original length of the caption
Private I As Integer        'used as a counter
Private Direction As String


Private Sub Form_Activate()
 marq1.Caption = "LSIK COMMUNITY"
        marq1.Alignment = 1 'make the alignment of the caption left-aligned
        orgcaption = marq1.Caption    'store the original caption entered above
        captionlen = Len(marq1.Caption)   'store the length of the original caption
        Direction = "out"   'set the direction to 'out'
        I = 1   'set the counter to 1
        Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
If Direction = "out" Then
        'the following line drops one character every time the timer interval occurs
        marq1.Caption = Mid(orgcaption, I)
        I = I + 1
        'if the counter is more than the length of the caption
        'change the direction so that the caption appears to
        'come in from the right of the caption box
        'to do this we need to set the direction to 'in' and the
        'caption alignment to 'right'
        If I > captionlen Then
            Direction = "in"
            marq1.Alignment = 2  'set caption alignment to right-aligned
            I = 1 'reset counter
        End If
    Else 'if direction is 'in'
        'the following line increases the visible caption by one character
        marq1.Caption = Left(orgcaption, I)
        I = I + 1
        'if the counter exceeds the number of characters that
        'are visible in the caption box we need to change the
        'direction to 'out' and the alignment to 'left-aligned'
        'NOTE: to find out how many characters are visible
        'in a caption box, physically count them when the caption box
        'is first visible i.e. before the scrolling starts.
        'The number of characters visible in a caption box depends
        'on the size of the caption box, the size of the font and
        'the font choosen.
        If I > 25 Then
            Direction = "out"
            marq1.Alignment = 2 'set the alignment to 'left-aligned'
            I = 1 'reset counter
        End If
    End If
End Sub



























Private Sub Command2_Click()


End Sub


Private Sub Command3_Click()

End Sub

Private Sub btnexit_Click()
LoginSukses = False

End
End Sub

Private Sub btnlogin_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("Isilah username dan password"), vbOKOnly + vbInformation
Text1.SetFocus
Exit Sub
End If

If Adodc1.Recordset.RecordCount = 0 Then
MsgBox ("Username dan password anda belum terdaftar"), vbOKOnly + vbCritical
Exit Sub
End If

Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
If Text1 = Adodc1.Recordset.Fields("Username") Then
If Text2 = Adodc1.Recordset.Fields("Password") Then
LoginSukses = True

Unload Me
MDI.Label1(0).Visible = True


End If
End If
Adodc1.Recordset.MoveNext
Loop
If Not LoginSukses Then
MsgBox ("Username atau password anda salah"), vbOKOnly + vbCritical, "Pesan Kesalahan"
SendKeys "{Home}+{End}"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus

End If
End Sub

Private Sub Delete_Click()
With Adodc1.Recordset
.Delete
End With
End Sub

Private Sub SignUp_Click()
With Adodc1.Recordset
.AddNew
!UserName = Text1.Text
!Password = Text2.Text
.Update
End With
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Time
End Sub
