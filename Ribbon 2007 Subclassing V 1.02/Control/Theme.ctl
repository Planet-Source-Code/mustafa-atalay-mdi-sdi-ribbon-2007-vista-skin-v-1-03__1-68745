VERSION 5.00
Begin VB.UserControl Tema 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   ControlContainer=   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   Begin VB.PictureBox ArkaPlan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   2340
      Picture         =   "Theme.ctx":0000
      ScaleHeight     =   2220
      ScaleWidth      =   3390
      TabIndex        =   14
      Top             =   450
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.PictureBox Tema7 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   1  'Dash
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   60
      ScaleWidth      =   7005
      TabIndex        =   0
      Top             =   4455
      Width           =   7005
   End
   Begin VB.PictureBox Tema4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   7035
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   60
      TabIndex        =   11
      Top             =   450
      Width           =   60
   End
   Begin VB.PictureBox Btns 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5370
      Picture         =   "Theme.ctx":0E46
      ScaleHeight     =   225
      ScaleWidth      =   420
      TabIndex        =   10
      Top             =   945
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Btns 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   5820
      Picture         =   "Theme.ctx":1374
      ScaleHeight     =   225
      ScaleWidth      =   390
      TabIndex        =   9
      Top             =   945
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox Btns 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   6240
      Picture         =   "Theme.ctx":1866
      ScaleHeight     =   225
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   945
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Btn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5430
      Picture         =   "Theme.ctx":1FEC
      ScaleHeight     =   225
      ScaleWidth      =   420
      TabIndex        =   7
      Tag             =   "Pasif"
      ToolTipText     =   "Pencereyi Simge Durmuna Getir"
      Top             =   120
      Width           =   420
   End
   Begin VB.PictureBox Btn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   5880
      Picture         =   "Theme.ctx":251A
      ScaleHeight     =   225
      ScaleWidth      =   390
      TabIndex        =   6
      Tag             =   "Pasif"
      Top             =   120
      Width           =   390
   End
   Begin VB.PictureBox Btn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   6300
      Picture         =   "Theme.ctx":2A0C
      ScaleHeight     =   225
      ScaleWidth      =   615
      TabIndex        =   5
      Tag             =   "Pasif"
      ToolTipText     =   "Pencereyi Kapat"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Btns 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   6240
      Picture         =   "Theme.ctx":3192
      ScaleHeight     =   225
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   645
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Btns 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   5820
      Picture         =   "Theme.ctx":3918
      ScaleHeight     =   225
      ScaleWidth      =   390
      TabIndex        =   3
      Top             =   645
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox Btns 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   5370
      Picture         =   "Theme.ctx":3E0A
      ScaleHeight     =   225
      ScaleWidth      =   420
      TabIndex        =   2
      Top             =   645
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Tema5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   1  'Dash
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6990
      Left            =   6945
      Negotiate       =   -1  'True
      ScaleHeight     =   6990
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   450
      Width           =   60
   End
   Begin VB.PictureBox Tema2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "Theme.ctx":4338
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   467
      TabIndex        =   12
      Top             =   0
      Width           =   7005
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   495
         TabIndex        =   13
         Top             =   105
         Width           =   720
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   90
         Stretch         =   -1  'True
         Top             =   45
         Width           =   300
      End
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2925
      Picture         =   "Theme.ctx":CCD6
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Menue 
      Height          =   3000
      Left            =   45
      Picture         =   "Theme.ctx":D118
      Top             =   450
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Menu Baslik 
      Caption         =   "Kapat"
      Begin VB.Menu MENU 
         Caption         =   "Kapat"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Tema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================================
'Product Version        : ActiveX Ribbon 2007 Vista Skin v 1.01
'Codec By and Source    : CoMpuTerBoY - Mustafa ATALAY
'From                   : Türkiye - Ottoman
'Mail-Msn               : sofi@odtu.com
'General Technics       : Subclassing , WindowPos
'=============================================================================================

Option Explicit

Private TemaAktif                    As Boolean
Private UstAktif                     As Boolean

Private KapatButton                  As Boolean
Private TamEkranButton               As Boolean
Private SimgeDurumuButton            As Boolean
Private YardimButton                 As Boolean


Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                              As Long
    dwFlags                             As TRACKMOUSEEVENT_FLAGS
    hwndTrack                           As Long
    dwHoverTime                         As Long
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Integer


Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const HWND_TOP = 0
Private Const WM_NCACTIVATE = &H86

Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    hwnd_notopmost = -2
End Enum

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = -20
Private Const SM_CXFRAME = 32
Private Const SM_CYCAPTION = 4
Private Const SM_CXDLGFRAME = 7
Private m_hForm As Long
Private m_Active As Boolean

Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" ( _
    ByVal pszThemeFileName As Long, _
    ByVal dwMaxNameChars As Long, _
    ByVal pszColorBuff As Long, _
    ByVal cchMaxColorChars As Long, _
    ByVal pszSizeBuff As Long, _
    ByVal cchMaxSizeChars As Long _
   ) As Long

Private Const THEME_BLUE = 1
Private Const THEME_OLIVE = 2
Private Const THEME_SILVER = 3

Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = 3  ' MSG_AFTER Or MSG_before                                   'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type
Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Sub Image1_DblClick()
Unload UserControl.Parent
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl.Parent.ZOrder
Resize True, True, True
PopupMenu Baslik, 0, Image1.left - Image1.left + 4, Image1.top - Image1.top + 30
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    
    With PropBag
        KapatButton = .ReadProperty("ButtonKapat", True)
        TamEkranButton = .ReadProperty("ButtonTamEkran", True)
        SimgeDurumuButton = .ReadProperty("ButtonSimgeDurumu", True)
        YardimButton = .ReadProperty("ButtonYardim", True)
    End With
    
    Dim retval As Long
    If Ambient.UserMode Then
        m_hForm = UserControl.Parent.hwnd
        Call Subclass_Start(m_hForm)
        
        Call Subclass_AddMsg(m_hForm, WM_SYSCOMMAND, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_MOVING, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_LBUTTONDOWN, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_SIZE, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_SHOWWINDOW, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_SETFOCUS, MSG_AFTER)
        'Call Subclass_AddMsg(m_hForm, WM_ACTIVATE, MSG_after)
        Call Subclass_AddMsg(m_hForm, MSM_NCACTIVATE, MSG_BEFORE)
        Call Subclass_AddMsg(m_hForm, WM_NCLBUTTONDOWN, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_PAINT, MSG_AFTER)
        Call Subclass_AddMsg(m_hForm, WM_ACTIVATEAPP, MSG_AFTER)
        
        Dim I
        For I = 1 To Btn.Count
        Call Subclass_Start(Btn(I).hwnd)
        Call Subclass_AddMsg(Btn(I).hwnd, WM_MOUSELEAVE, MSG_AFTER)
        Next
            
    End If
    
    
End Sub

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
    Dim X As Long
    Dim Y As Long
    Dim a As Boolean

    Select Case lng_hWnd
    Case m_hForm
        Select Case uMsg
        
        Case WM_NCLBUTTONDOWN
                Select Case wParam
                    Case HTRIGHT
                        Resize True, True, True, SWP_NOZORDER
                        Debug.Print "Sað Kýsým"
                    Case HTBOTTOM
                        Resize True, True, True, SWP_NOZORDER
                        Debug.Print "Alt Kýsým"
                    Case HTBOTTOMRIGHT
                        Resize True, True, True, SWP_NOZORDER
                        Debug.Print "Alt Orta Kýsým"
                End Select
        Case WM_SYSCOMMAND
            Select Case wParam
            Case SC_CLOSE:      Debug.Print "Kapandim..."
                Subclass_StopAll
            Case SC_RESTORE:    Debug.Print "Küçüldüm..."
                UserControl.Parent.WindowState = 0
            Case SC_MAXIMIZE:   Debug.Print "Büyüdüm..."
                UserControl.Parent.WindowState = 2
                If Screen.ActiveForm Is Nothing Then Exit Sub
                If Screen.ActiveForm.WindowState = 2 Then Screen.ActiveForm.WindowState = 2
            Case SC_MOVE:       Debug.Print "Taþýnýyorum..."
                
                If UserControl.Parent.WindowState = 2 Then Exit Sub
                UserControl.Parent.ZOrder
                ReleaseCapture
                SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, &HF012&, 0&
                UstAktif = True
                
            Case SC_MINIMIZE:   Debug.Print "Minimize Oluyom..."
                If Screen.ActiveForm Is Nothing Then Exit Sub
                If Screen.ActiveForm.WindowState = 2 Then Screen.ActiveForm.WindowState = 2
            Case Else
                Call Resize(True, True, True)
                
            End Select
        Case WM_SIZE, WM_MOVING, WM_LBUTTONDOWN
            Select Case wParam
            Case MK_LBUTTON
                Debug.Print "Sol Týk..."
            Case Else
            End Select
                If UstAktif = True Then UserControl.Parent.ZOrder
                Call Resize(True, True, True)
                UstAktif = False
                
        Case WM_ACTIVATEAPP
            Select Case wParam
               Case WA_ACTIVE
                   Tema 1
                   UserControl.Parent.ZOrder
                   Call Resize(True, True, True)
                   Debug.Print "WA_ACTIVE"
               Case WA_CLICKACTIVE
                   Tema 1
                   UserControl.Parent.ZOrder
                   Call Resize(True, True, True)
                   Debug.Print "WA_CLICKACTIVE"
               Case WA_INACTIVE
                   Tema 2
                   Debug.Print "WA_INACTIVE"
            End Select
            
        Case MSM_NCACTIVATE
        
            Select Case wParam
               Case WA_ACTIVE
                   Tema 1
                   UserControl.Parent.ZOrder
                   Call Resize(True, True, True)
                   Debug.Print "ChildForm Active"
               Case WA_CLICKACTIVE
                   Tema 1
                   UserControl.Parent.ZOrder
                   Call Resize(True, True, True)
                   Debug.Print "ChildForm ClickActive"
               Case WA_INACTIVE
                   Tema 2
                   Debug.Print "ChildForm Passive"
            End Select
            

        Case WM_SETFOCUS
            Tema 1

            Call Resize(True, True, True)
            Debug.Print "Set Focus Oldum..."
            
        Case WM_SHOWWINDOW
                If lParam = 0 And wParam = 0 Then
                   Call SetParent(UserControl.hwnd, m_hForm)
                   Debug.Print "lParam = 0 And wParam = 0"
                ElseIf lParam = 0 And wParam = 1 Then
                   Call SetWindowLong(UserControl.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)
                   Call SetParent(UserControl.hwnd, GetParent(m_hForm))
                   Debug.Print "lParam = 0 And wParam = 1"
                End If
      End Select
    Case Btn(1).hwnd
    Debug.Print wParam & " - " & lParam
        If uMsg = WM_MOUSELEAVE Then
            Btn(1).Picture = Btns(4).Picture
            Btn(1).Tag = "Pasif"
        End If
    Case Btn(2).hwnd
        If uMsg = WM_MOUSELEAVE Then
            Btn(2).Picture = Btns(5).Picture
            Btn(2).Tag = "Pasif"
        End If
    Case Btn(3).hwnd
        If uMsg = WM_MOUSELEAVE Then
         Btn(3).Picture = Btns(6).Picture
            Btn(3).Tag = "Pasif"
        End If
    End Select
'Debug.Print uMsg

End Sub

Private Sub Btn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Select Case Index
Case 1
'On Error Resume Next
Call Subclass_StopAll
SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0&
Case 2
If UserControl.Parent.WindowState = 0 Then
SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0&
Btn(2).ToolTipText = "Pencereyi Önceki Boyutuna Getir"
Else
SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_RESTORE, 0&
Btn(2).ToolTipText = "Pencereyi Tam Ekran Yap"
Call Resize(True, True, True)
End If
Case 3
If UserControl.Parent.WindowState = 1 Then
UserControl.Parent.WindowState = 0
Else
UserControl.Parent.WindowState = 1
End If
End Select
End If
End Sub

Private Sub Btn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error GoTo Errs
    Dim tme As TRACKMOUSEEVENT_STRUCT
    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_LEAVE
        .hwndTrack = Btn(Index).hwnd
        If Btn(Index).Tag = "Pasif" Then
        Btn(Index).Picture = Btns(Index).Picture
        Btn(Index).Tag = "Aktif"
        End If
    End With
    Call TrackMouseEvent(tme)
Errs:
If UserControl.Parent.WindowState = 2 Then
Btn(2).ToolTipText = "Pencereyi Önceki Boyutuna Getir"
ElseIf UserControl.Parent.WindowState = 0 Then
Btn(2).ToolTipText = "Pencereyi Tam Ekran Yap"
End If
End Sub



Private Sub Label2_DblClick()
If TamEkranButton = False Then Exit Sub
If UserControl.Parent.WindowState = 0 Then
UserControl.Parent.WindowState = 2
Else
UserControl.Parent.WindowState = 0
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_MOVE, 0&
End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl.Parent.ZOrder
Resize True, True, True
End Sub

Private Sub Tema2_DblClick()

If TamEkranButton = False Then Exit Sub
If UserControl.Parent.WindowState = 0 Then
SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0&
Else
SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_RESTORE, 0&
End If

End Sub

Private Sub Tema2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
'If TypeOf UserControl.Parent Is MDIForm Then Exit Sub
SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_MOVE, 0&
End If
End Sub

Private Sub Tema2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl.Parent.ZOrder
Resize True, True, True
End Sub

Private Sub Tema5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
UserControl.Parent.ZOrder
Resize True, True, True
ReleaseCapture
SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
If Screen.ActiveForm Is Nothing Then Exit Sub
If Screen.ActiveForm.WindowState = 2 Then Screen.ActiveForm.WindowState = 2
End If
End Sub

Private Sub Tema5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If UserControl.Parent.WindowState = 1 Then
Tema5.MousePointer = 0
Exit Sub
End If
Tema5.MousePointer = 9
End Sub

Private Sub Tema7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
UserControl.Parent.ZOrder
Resize True, True, True
Select Case X
    Case Is > (UserControl.Parent.Width) - (10 * Screen.TwipsPerPixelX)
        ReleaseCapture
        SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
        Resize True, True, True
    Case Is < (10 * Screen.TwipsPerPixelX)
        ReleaseCapture
        SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, ByVal 0&
        Resize True, True, True
    Case Else
        ReleaseCapture
        SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOM, ByVal 0&
        Resize True, True, True
End Select
End If
End Sub

Private Sub Tema7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If UserControl.Parent.WindowState = 1 Then
Tema7.MousePointer = 0
Exit Sub
End If
Select Case X
    Case Is > (UserControl.Parent.Width) - (10 * Screen.TwipsPerPixelX)
        Tema7.MousePointer = 8
    Case Else
        Tema7.MousePointer = 7
End Select
End Sub



Public Sub Resize(Boyut As Boolean, Kýrp As Boolean, Pozisyon As Boolean, Optional lFlag As Long = SWP_FRAMECHANGED)
Dim cy, lStyle, cx
On Error Resume Next

If Ambient.UserMode Then
If Boyut = True Then

Dim r As RECT
GetWindowRect m_hForm, r


    cy = GetSystemMetrics(SM_CYCAPTION)
    cx = GetSystemMetrics(SM_CXFRAME)
    
    'Debug.Print CY & "-" & CX
    
    Select Case cy
        Case 20
        
        Case 19
        Tema2.Height = cy + cx
        Label2.top = 8 - cx
        Image1.top = 3 - 2
        Btn(1).top = 8 - 3
        Btn(2).top = 8 - 3
        Btn(3).top = 8 - 3
        Tema4.top = 23
        Tema5.top = 23
        Case 26
        Tema2.Height = cy + cx
        Label2.top = 8
        Image1.top = 3
        Btn(1).top = 8
        Btn(2).top = 8
        Btn(3).top = 8
        Tema4.top = 30
        Tema5.top = 30
    End Select
    


If TamEkranButton = False Then
Btn(1).left = r.right - r.left - Btn(1).Width - 4
Btn(2).left = r.right - r.left - Btn(2).Width - 4 - Btn(1).Width
Btn(3).left = r.right - r.left - Btn(3).Width + Btn(2).Width - 4 - Btn(2).Width - Btn(1).Width
Else
Btn(1).left = r.right - r.left - Btn(1).Width - 4
Btn(2).left = r.right - r.left - Btn(2).Width - 4 - Btn(1).Width
Btn(3).left = r.right - r.left - Btn(3).Width - 4 - Btn(2).Width - Btn(1).Width
End If


Tema5.left = r.right - r.left - cx
Tema2.Width = r.right - r.left
Tema7.Width = r.right - r.left
Tema4.Height = r.bottom - r.top - cx
Tema5.Height = r.bottom - r.top - cx
Tema7.top = r.bottom - r.top - cx

End If

If Kýrp = True Then
Dim lret As Long
Dim GRET As Long

' form görünümü genel kesim
On Error Resume Next
lret = CreateRoundRectRgn(4, cy + 4, UserControl.Parent.Width / 15 - 3, UserControl.Parent.Height / 15 - 3, 0, 0)
Call SetWindowRgn(UserControl.Parent.hwnd, lret, True)
Call DeleteObject(lret)

End If
If Pozisyon = True Then
Call SetWindowPos(UserControl.hwnd, HWND_TOP, UserControl.Parent.left / 15, UserControl.Parent.top / 15, UserControl.Parent.Width / 15, UserControl.Parent.Height / 15, lFlag)
End If
End If

    
End Sub


Private Sub UserControl_Show()

Image1.Picture = UserControl.Parent.Icon
Label2.Caption = UserControl.Parent.Caption

If KapatButton = False Then Btn(1).Visible = False
If TamEkranButton = False Then Btn(2).Visible = False
If SimgeDurumuButton = False Then Btn(3).Visible = False

Dim GRET As Long
On Error Resume Next
GRET = CreateRoundRectRgn(0, 0, UserControl.Width / 15, UserControl.Height / 15, 30, 30)
Call SetWindowRgn(UserControl.hwnd, GRET, True)
Call DeleteObject(GRET)

 UserControl.Parent.Picture = ArkaPlan.Picture
 UserControl.Parent.BackColor = RGB(70, 70, 70)

MenuYap
End Sub

Private Sub UserControl_Terminate()
    'On Error GoTo Errs
    'If Ambient.UserMode Then
    Call Subclass_StopAll

    'End If
Errs:
End Sub

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'On Error GoTo Errs
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'On Error GoTo Errs
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'On Error GoTo Errs
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
  If aBuf(1) = 0 Then
  
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

    I = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, I, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

Private Sub Subclass_StopAll()
'On Error GoTo Errs
  Dim I As Long
  
  If Not sc_aSubData Then Exit Sub
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    I = I - 1                                                                           'Next element
  Loop
Errs:
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'On Error GoTo Errs
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'On Error GoTo Errs
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
    If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If
Errs:

End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

Private Sub Tema(No)
Select Case No
Case 1 'Aktif
Tema4.BackColor = &H404040
Tema5.BackColor = &H404040
Tema7.BackColor = &H404040
TemaAktif = True
Case 2 'Pasif
Tema4.BackColor = &H808080
Tema5.BackColor = &H808080
Tema7.BackColor = &H808080
TemaAktif = False
End Select
End Sub

Private Sub MENU_Click(Index As Integer)
Unload UserControl.Parent
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "ButtonKapat", KapatButton, True
        .WriteProperty "ButtonTamEkran", TamEkranButton, True
        .WriteProperty "ButtonSimgeDurumu", SimgeDurumuButton, True
        .WriteProperty "ButtonYardim", YardimButton, True
    End With

End Sub

Public Property Get ButtonKapat() As Boolean
    ButtonKapat = KapatButton
End Property

Public Property Let ButtonKapat(ByVal vData As Boolean)
KapatButton = vData
PropertyChanged "ButtonKapat"
End Property

Public Property Get ButtonTamEkran() As Boolean
    ButtonTamEkran = TamEkranButton
End Property

Public Property Let ButtonTamEkran(ByVal vData As Boolean)
TamEkranButton = vData
PropertyChanged "ButtonTamEkran"
End Property

Public Property Get ButtonSimgeDurumu() As Boolean
    ButtonSimgeDurumu = SimgeDurumuButton
End Property

Public Property Let ButtonSimgeDurumu(ByVal vData As Boolean)
SimgeDurumuButton = vData
PropertyChanged "ButtonSimgeDurumu"
End Property

Private Sub MenuYap()
    startODMenus UserControl.Parent, True
    With CustomMenu
        Set .Picture = Menue.Picture
        .PosX = 30
        .Icon.Add Image2.Picture, "Yeni"
        .FontBold = True
        .Texture = True
    End With
    MenuMode = XPlook
    With CustomColor
        .ForeColor = vbWhite
        .BackColor = vbYellow
        .BorderColor = vbWhite
        .DefTextColor = vbRed
        .ForeColor = vbYellow
        .HilightColor = RGB(40, 40, 40)
        .SelectedTextColor = vbRed
        .MenuTextColor = vbRed
        .NormalColor = vbRed
        .RECTColor = vbWhite
    End With
End Sub





