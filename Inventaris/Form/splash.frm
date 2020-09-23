VERSION 5.00
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   ControlBox      =   0   'False
   Icon            =   "splash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splash.frx":0E42
   ScaleHeight     =   3735
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   360
      Top             =   2640
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "www.lsik.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2880
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   3390
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      Caption         =   "For more information visit"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3420
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Copyright © 2006 LSIK Community. All right reserved"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3170
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Label1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   5295
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()
Shell ("explorer " & "http://www.lsik.net")
End Sub


Private Sub Timer1_Timer()
Unload Me
MDI.Show
End Sub
