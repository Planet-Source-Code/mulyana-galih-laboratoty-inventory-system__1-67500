VERSION 5.00
Begin VB.Form Rekomendasi 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requirement System"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Rekomendasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "&Close"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      cfore           =   -2147483647
      font            =   "Rekomendasi.frx":74F2
      mode            =   0
      value           =   0   'False
      image           =   "Rekomendasi.frx":751E
      cfhover         =   -2147483647
      cback           =   -2147483639
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "sudah terinstalasi VB 6.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "1 set komputer yang masih nyala dan bagus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "window Operating System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Makasimum intel pentium IV 3.6 GB Duo Core"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum intel pentium II  200 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistem Rekomendasi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Rekomendasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvButtons_H1_Click()
Unload Me
End Sub
