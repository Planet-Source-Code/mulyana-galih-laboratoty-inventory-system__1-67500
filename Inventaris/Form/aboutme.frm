VERSION 5.00
Begin VB.Form aboutme 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "about me"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      cFore           =   -2147483643
      cFHover         =   -2147483643
      cBhover         =   -2147483629
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "aboutme.frx":0000
      cBack           =   -2147483632
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Born Cianjur, January 11th 1983. and now live in Jakarta. i hope u give me a critic, via post email "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "about me:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2850
      Left            =   2280
      Picture         =   "aboutme.frx":0480
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2400
   End
End
Attribute VB_Name = "aboutme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Private Sub Form_Activate()
Unload REPORT_PRLTN
Unload REPORT_KOMP
Unload REPORT_FAS
Unload REPORT_ATK
Unload ADMINS
Unload INPUTKOMPUTER
Unload FASILITAS
Unload PERALATAN
Unload ATK
End Sub

Private Sub Form_Load()
MakeTransparent aboutme
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub
