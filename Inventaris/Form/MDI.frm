VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI.frx":0E42
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1800
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   0
      ScaleHeight     =   163.022
      ScaleMode       =   0  'User
      ScaleWidth      =   381.89
      TabIndex        =   0
      Top             =   600
      Width           =   1485
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1080
         Top             =   720
      End
      Begin MSComctlLib.ListView lvMenu 
         Height          =   6495
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   11456
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000A&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   4935
         Left            =   0
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000A&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000005&
         BorderStyle     =   2  'Dash
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   2655
         Left            =   -120
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   31
      MaskColor       =   16744576
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":664D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6E79
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":72AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7766
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7FB7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4080
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":83E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":9D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B708
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":D09A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":D976
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":E652
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":EF36
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":F812
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Index           =   0
      Left            =   2880
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":111A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":12B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":144CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":14DAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn1"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn2"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn3"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn4"
            Object.ToolTipText     =   "Report"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin Project1.lvButtons_H lvButtons_H1 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   7050
         TabIndex        =   4
         ToolTipText     =   "www.lsik.net"
         Top             =   30
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   873
         Caption         =   "LABORATORIUM SISTEM INFORMASI DAN KEPUTUSAN"
         CapAlign        =   1
         BackStyle       =   3
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483644
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   -2147483644
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483648
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnulog 
         Caption         =   "&Logout"
      End
      Begin VB.Menu mnusign 
         Caption         =   "&SignUp"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print Report"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuinvent 
      Caption         =   "&Inventory"
      Begin VB.Menu mnukomputer 
         Caption         =   "&Komputer"
         Begin VB.Menu mnuinput 
            Caption         =   "&Input"
         End
         Begin VB.Menu mnuedit 
            Caption         =   "&Edit"
         End
         Begin VB.Menu mnuhapus 
            Caption         =   "&Hapus"
         End
      End
      Begin VB.Menu mnuatk 
         Caption         =   "&ATK"
         Begin VB.Menu mnuinput2 
            Caption         =   "&Input"
         End
         Begin VB.Menu mnuedit2 
            Caption         =   "&Edit"
         End
         Begin VB.Menu mnuhapus2 
            Caption         =   "&Hapus"
         End
      End
      Begin VB.Menu mnuperalatan 
         Caption         =   "&Peralatan"
         Begin VB.Menu mnuinput3 
            Caption         =   "&Input"
         End
         Begin VB.Menu mnuedit3 
            Caption         =   "&Edit"
         End
         Begin VB.Menu mnuhapus3 
            Caption         =   "&Hapus"
         End
      End
      Begin VB.Menu mnufasilitas 
         Caption         =   "&Fasilitas"
         Begin VB.Menu mnuinput4 
            Caption         =   "&Input"
         End
         Begin VB.Menu mnuedit4 
            Caption         =   "&Edit"
         End
         Begin VB.Menu mnuhapus4 
            Caption         =   "&Hapus"
         End
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
      Begin VB.Menu mnureportkomp 
         Caption         =   "&Komputer"
      End
      Begin VB.Menu mnureportatk 
         Caption         =   "&Alat Tulis Kantor"
      End
      Begin VB.Menu mnureportfasilitas 
         Caption         =   "&Fasilitas"
      End
      Begin VB.Menu mnureportprltn 
         Caption         =   "&Peralatan"
      End
   End
   Begin VB.Menu mnuwindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuhelp0 
      Caption         =   "&Help"
      Begin VB.Menu mnurequirement 
         Caption         =   "&Requirement"
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "about &me"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Judul As String = "LSIK Inventory System - "


Private Sub lvButtons_H1_Click()
Shell ("explorer " & "http://www.lsik.net")
End Sub

Private Sub lvMenu_DblClick()
 Select Case lvMenu.SelectedItem.Key
        'For Sales
        Case "user":  ADMINS.Show
        Unload INPUTKOMPUTER
        Unload PERALATAN
        Unload ATK
        Unload FASILITAS
        ADMINS.Frame5.Visible = True
        
        
        Case "komputer": INPUTKOMPUTER.Show
        Unload ADMINS
        Unload PERALATAN
        Unload ATK
        Unload FASILITAS
        INPUTKOMPUTER.Frame5.Visible = True
        INPUTKOMPUTER.Frame2.Visible = False
         
        Case "peralatan":  PERALATAN.Show
        Unload INPUTKOMPUTER
        Unload ADMINS
        Unload ATK
        Unload FASILITAS
        PERALATAN.Frame5.Visible = True
                
        Case "atk": ATK.Show
        Unload ADMINS
        Unload INPUTKOMPUTER
        Unload PERALATAN
        Unload FASILITAS
        ATK.Frame5.Visible = True
        
        Case "fasilitas":  FASILITAS.Show
        Unload INPUTKOMPUTER
        Unload ADMINS
        Unload PERALATAN
        Unload ATK
        Unload FASILITAS
        FASILITAS.Frame5.Visible = True
               
        Case "laporan": REPORTING.Show
        
    End Select

End Sub



Private Sub MDIForm_Load()
Me.Show
Me.Caption = Judul + CStr(Now) + ""

Stat = False

With lvMenu
        Set .SmallIcons = ImageList2
        Set .Icons = ImageList2
        'For Sales
        .ListItems.Add , "user", "ADMIN", 1, 1
        .ListItems.Add , "komputer", "Komputer Laboratorium", 2, 2
        .ListItems.Add , "peralatan", "Peralatan Laboratorium", 3, 3
        .ListItems.Add , "atk", "Alat Tulis Kantor", 4, 4
        .ListItems.Add , "fasilitas", "Fasilitas Laboratorium", 5, 5
        .ListItems.Add , "laporan", "View Report", 6, 6
    End With
Label1(0).Visible = False
LOGIN.Show vbModal
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuabout_Click()
aboutme.Show vbModal
End Sub

Private Sub mnuedit_Click()
INPUTKOMPUTER.Show
INPUTKOMPUTER.Frame2.Visible = True
INPUTKOMPUTER.Frame5.Visible = False
INPUTKOMPUTER.Frame3.Visible = False
INPUTKOMPUTER.Frame1.Visible = False
End Sub

Private Sub mnuedit2_Click()
ATK.Show
ATK.Frame2.Visible = True
ATK.Frame5.Visible = False
ATK.Frame3.Visible = False
ATK.Frame1.Visible = False
End Sub

Private Sub mnuedit3_Click()
PERALATAN.Show
PERALATAN.Frame3.Visible = False
PERALATAN.Frame5.Visible = False
PERALATAN.Frame2.Visible = True
PERALATAN.Frame1.Visible = False
End Sub

Private Sub mnuedit4_Click()
FASILITAS.Show
FASILITAS.Frame2.Visible = True
FASILITAS.Frame5.Visible = False
FASILITAS.Frame1.Visible = False
FASILITAS.Frame3.Visible = False
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuhapus_Click()
INPUTKOMPUTER.Show
INPUTKOMPUTER.Frame2.Visible = False
INPUTKOMPUTER.Frame5.Visible = False
INPUTKOMPUTER.Frame3.Visible = False
INPUTKOMPUTER.Frame1.Visible = True
End Sub

Private Sub mnuhapus2_Click()
ATK.Show
ATK.Frame2.Visible = False
ATK.Frame5.Visible = False
ATK.Frame3.Visible = False
ATK.Frame1.Visible = True
End Sub

Private Sub mnuhapus3_Click()
PERALATAN.Show
PERALATAN.Frame1.Visible = True
PERALATAN.Frame3.Visible = False
PERALATAN.Frame5.Visible = False
PERALATAN.Frame2.Visible = False

End Sub

Private Sub mnuhapus4_Click()
FASILITAS.Show
FASILITAS.Frame2.Visible = False
FASILITAS.Frame5.Visible = False
FASILITAS.Frame1.Visible = True
FASILITAS.Frame3.Visible = False
End Sub

Private Sub mnuhelp_Click()
Shell ("explorer mailto:" & "mulyanagalih@gmail.com")
End Sub

Private Sub mnuinput_Click()
INPUTKOMPUTER.Show
INPUTKOMPUTER.Frame2.Visible = False
INPUTKOMPUTER.Frame5.Visible = False
INPUTKOMPUTER.Frame3.Visible = True
INPUTKOMPUTER.Frame1.Visible = False
End Sub

Private Sub mnuinput2_Click()
ATK.Show
ATK.Frame2.Visible = False
ATK.Frame5.Visible = False
ATK.Frame3.Visible = True
ATK.Frame1.Visible = False
End Sub

Private Sub mnuinput3_Click()
PERALATAN.Show
PERALATAN.Frame3.Visible = True
PERALATAN.Frame5.Visible = False
PERALATAN.Frame2.Visible = False
PERALATAN.Frame1.Visible = False
End Sub

Private Sub mnuinput4_Click()
FASILITAS.Show
FASILITAS.Frame3.Visible = True
FASILITAS.Frame5.Visible = False
End Sub

Private Sub mnulog_Click()
Unload REPORT_PRLTN
Unload REPORT_KOMP
Unload REPORT_FAS
Unload REPORT_ATK
Unload ADMINS
Unload INPUTKOMPUTER
Unload FASILITAS
Unload PERALATAN
Unload ATK
LOGIN.Show vbModal
End Sub

Private Sub mnuprint_Click()
cetak.Show vbModal
End Sub

Private Sub mnureportatk_Click()
REPORT_ATK.Show
Unload REPORT_PRLTN
Unload REPORT_KOMP
Unload REPORT_FAS
End Sub

Private Sub mnureportfasilitas_Click()
REPORT_FAS.Show
Unload REPORT_PRLTN
Unload REPORT_ATK
Unload REPORT_KOMP
End Sub

Private Sub mnureportkomp_Click()
REPORT_KOMP.Show
Unload REPORT_PRLTN
Unload REPORT_ATK
Unload REPORT_FAS
End Sub

Private Sub mnureportprltn_Click()
REPORT_PRLTN.Show
Unload REPORT_KOMP
Unload REPORT_ATK
Unload REPORT_FAS
End Sub

Private Sub mnurequirement_Click()
Rekomendasi.Show vbModal
End Sub

Private Sub mnusign_Click()
ADMINS.Show
FASILITAS.Frame2.Visible = False
FASILITAS.Frame5.Visible = False
FASILITAS.Frame1.Visible = False
FASILITAS.Frame3.Visible = True
End Sub


Private Sub Timer1_Timer()
Label1(0).Caption = Time
End Sub

Private Sub Timer2_Timer()
If Stat = False Then
    Me.Caption = Judul + CStr(Now) + ""
Else
    Me.Caption = Judul + CStr(Now) + " - " + User
End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Btn As String

Btn = Button.Key

Select Case Btn
    Case "btn1"
       ATK.Adodc1.Refresh
       INPUTKOMPUTER.Adodc1.Refresh
       PERALATAN.Adodc1.Refresh
       FASILITAS.Adodc1.Refresh
       ADMINS.Adodc1.Refresh
    Case "btn2"
        cetak.Show vbModal
    Case "btn3"
        Shell ("explorer mailto:" & "mulyanagalih@gmail.com")
    Case "btn4"
       REPORTING.Show vbModal
    
    End Select
End Sub

