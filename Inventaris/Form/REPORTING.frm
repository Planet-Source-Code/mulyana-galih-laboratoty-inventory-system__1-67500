VERSION 5.00
Begin VB.Form REPORTING 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporting Option"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   Icon            =   "REPORTING.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Reporting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000009&
         Caption         =   "Report Komputer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000009&
         Caption         =   "Report ATK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000009&
         Caption         =   "Report Fasilitas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H80000009&
         Caption         =   "Report Peralatan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin Project1.lvButtons_H tampil 
         Height          =   405
         Left            =   2280
         TabIndex        =   5
         Top             =   2160
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   714
         Caption         =   "&Tampilkan"
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
         Image           =   "REPORTING.frx":74F2
         cBack           =   16777215
      End
      Begin Project1.lvButtons_H close 
         Height          =   405
         Left            =   3720
         TabIndex        =   7
         Top             =   2160
         Width           =   945
         _ExtentX        =   1667
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
         ImgAlign        =   2
         Image           =   "REPORTING.frx":7A22
         cBack           =   16777215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Pilihlah opsih dibawah yang  akan ditampilkan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "REPORTING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub close_Click()
Unload Me
End Sub

Private Sub tampil_Click()
If Opt(0).Value = True Then
REPORT_KOMP.Show
Unload REPORT_PRLTN
Unload REPORT_ATK
Unload REPORT_FAS
ElseIf Opt(1).Value = True Then
REPORT_ATK.Show
Unload REPORT_PRLTN
Unload REPORT_KOMP
Unload REPORT_FAS
ElseIf Opt(2).Value = True Then
REPORT_FAS.Show
Unload REPORT_PRLTN
Unload REPORT_ATK
Unload REPORT_KOMP
ElseIf Opt(3).Value = True Then
REPORT_PRLTN.Show
Unload REPORT_KOMP
Unload REPORT_ATK
Unload REPORT_FAS
End If
Unload Me
End Sub
