VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Child 1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   9540
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1575
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   3015
   End
   Begin Project1.Tema Tema1 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ButtonTamEkran  =   0   'False
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   615
      Left            =   11400
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2175
      Left            =   10680
      TabIndex        =   0
      Top             =   2160
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

