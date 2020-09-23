VERSION 5.00
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00464646&
   Caption         =   "MDIForm"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7395
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   7365
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin Project1.Tema Tema1 
         Height          =   975
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1720
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Dosya"
      Index           =   0
      Begin VB.Menu yeni 
         Caption         =   "Yeni Oluþtur"
      End
      Begin VB.Menu AltMenu1 
         Caption         =   "Dosyadan Aç"
         Index           =   1
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu AltMen 
         Caption         =   "Kaydet"
         Index           =   2
      End
      Begin VB.Menu AltMe 
         Caption         =   "Farklý Kaydet"
         Index           =   3
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu as 
         Caption         =   "Sayfa Önizleme"
      End
      Begin VB.Menu df 
         Caption         =   "Yazdýr"
      End
      Begin VB.Menu sdfsddsf 
         Caption         =   "-"
      End
      Begin VB.Menu sad 
         Caption         =   "Hakkýnda"
      End
      Begin VB.Menu sdfsdfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu dsf 
         Caption         =   "Çýkýþ"
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Araçlar"
      Index           =   1
      Begin VB.Menu AltMenu2 
         Caption         =   "Baðlantý Ayarlarý"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
Form1.Show
Form2.Show
End Sub



