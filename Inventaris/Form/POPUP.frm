VERSION 5.00
Begin VB.Form POPUP 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuinput 
      Caption         =   "&Input"
      Begin VB.Menu mnuinputkom 
         Caption         =   "Input &Komputer"
      End
      Begin VB.Menu mnuinputatk 
         Caption         =   "Input &ATK"
      End
      Begin VB.Menu mnuinputfasilitas 
         Caption         =   "Input &Fasilitas"
      End
      Begin VB.Menu mnuinputperalatan 
         Caption         =   "Input &Peralatan"
      End
   End
   Begin VB.Menu mnukomputer 
      Caption         =   "&Komputer"
      Begin VB.Menu mnuinputk 
         Caption         =   "&Input"
      End
      Begin VB.Menu mnueditk 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuhapusk 
         Caption         =   "&Hapus"
      End
   End
   Begin VB.Menu mnuatk 
      Caption         =   "&ATK"
      Begin VB.Menu mnuinputatk1 
         Caption         =   "&Input"
      End
      Begin VB.Menu mnueditatk 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuhapusatk 
         Caption         =   "&Hapus"
      End
   End
   Begin VB.Menu mnufasilitas 
      Caption         =   "&Fasilitas"
      Begin VB.Menu mnuinputfasilitas1 
         Caption         =   "&Input"
      End
      Begin VB.Menu mnueditfasilitas 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuhapusfasilitas 
         Caption         =   "&Hapus"
      End
   End
   Begin VB.Menu mnuperalatan 
      Caption         =   "&Peralatan"
      Begin VB.Menu mnuinputperalatan1 
         Caption         =   "&Input"
      End
      Begin VB.Menu mnueditperalatan 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuhapusperalatan 
         Caption         =   "&Hapus"
      End
   End
End
Attribute VB_Name = "POPUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuinputkom_Click()
INPUTKOMPUTER.Show
End Sub
