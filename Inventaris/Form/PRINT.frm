VERSION 5.00
Begin VB.Form CETAK 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Option"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "PRINT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Print"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
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
         TabIndex        =   5
         Top             =   1800
         Width           =   2055
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
         TabIndex        =   4
         Top             =   1440
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
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
      Begin Project1.lvButtons_H cetak 
         Default         =   -1  'True
         Height          =   405
         Left            =   2640
         TabIndex        =   6
         Top             =   2160
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   714
         Caption         =   "&Cetak"
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
         Image           =   "PRINT.frx":74F2
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
         Image           =   "PRINT.frx":79A4
         cBack           =   16777215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Pilihlah opsih dibawah yang  akan dicetak"
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
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "CETAK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cetak_Click()
If Opt(0).Value = True Then
REPORT_KOMP.PrintReport (False)
ElseIf Opt(1).Value = True Then
REPORT_ATK.PrintReport (False)
ElseIf Opt(2).Value = True Then
REPORT_FAS.PrintReport False
ElseIf Opt(3).Value = True Then
REPORT_PRLTN.PrintReport False
End If
Unload Me
End Sub

Private Sub close_Click()
Unload Me
End Sub
