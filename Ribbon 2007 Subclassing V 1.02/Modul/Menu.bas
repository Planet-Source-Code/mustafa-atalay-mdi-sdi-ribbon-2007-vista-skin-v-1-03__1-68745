Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreatePatternBrush Lib "GDI32" (ByVal hBitmap As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "User32" () As Long
Public Declare Function GetMenu Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuInfo Lib "User32" (ByVal hWnd As Long, mInfo As MENUINFO) As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetWindowDC Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetMenuInfo Lib "User32" (ByVal hWnd As Long, mInfo As MENUINFO) As Long

Public Type MENUINFO
cbSize              As Long
fMask               As Long
dwStyle             As Long
cyMax               As Long
hbrBack             As Long
dwContextHelpID     As Long
dwMenuData          As Long
End Type

Public Const MIM_BACKGROUND = &H2

