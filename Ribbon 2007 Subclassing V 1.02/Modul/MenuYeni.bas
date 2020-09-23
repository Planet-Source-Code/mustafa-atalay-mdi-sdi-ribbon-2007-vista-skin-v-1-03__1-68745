Attribute VB_Name = "Menu"
'Original sourcecode by : G. D. Sever    aka "The Hand" (thehand@elitevb.com)
'Modify by: Kowshar Ahmed > bioscopedvd_ksnb@yahoo.com or versatileplayer@gmail.com
'web: http://www.bioscopedvd.com
Option Explicit
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, MI As MENUINFO) As Long

Private Enum MENUINFO_STYLES
MNS_NOCHECK = &H80000000
MNS_MODELESS = &H40000000
MNS_DRAGDROP = &H20000000
MNS_AUTODISMISS = &H10000000
MNS_NOTIFYBYPOS = &H8000000
MNS_CHECKORBMP = &H4000000
End Enum
Private Enum MENUINFO_MASKS
MIM_MAXHEIGHT = &H1
MIM_BACKGROUND = &H2
MIM_HELPID = &H4
MIM_MENUDATA = &H8
MIM_STYLE = &H10
MIM_APPLYTOSUBMENUS = &H80000000
End Enum
Private Type MENUINFO
cbSize As Long
fMask As MENUINFO_MASKS
dwStyle As MENUINFO_STYLES
cyMax As Long
hbrBack As Long
dwContextHelpID As Long
dwMenuData As Long
End Type

Public Type MenuSetting
    Texture                                   As Boolean
    Picture                                   As StdPicture
    Icon                                      As Collection
    'Font                                      As StdFont
    FontName                                  As String
    FontBold                                  As Boolean
    FontItalic                                As Boolean
    FontStrikeOut                             As Boolean
    FontUnderline                             As Boolean
    CustomFonts                               As Collection
    UseCustomFonts                            As Boolean
    PosX                                      As Long
    
End Type
Public CustomMenu                             As MenuSetting
Public Type ColorSetting
    HilightColor                              As OLE_COLOR
    NormalColor                               As OLE_COLOR
    ForeColor                                 As OLE_COLOR
    BackColor                                 As OLE_COLOR
    SelectedTextColor                         As OLE_COLOR
    MenuTextColor                             As OLE_COLOR
    BorderColor                               As OLE_COLOR
    DefTextColor                              As OLE_COLOR
    RECTColor                                 As OLE_COLOR
End Type
Public CustomColor                            As ColorSetting
Public Enum MenuLook
       VBLook = 0
       XPlook = 1
End Enum
#If False Then
Private VBLook, XPlook
#End If

Public MenuMode                               As MenuLook
' ********************************************************************************
'     Couple of things the module uses to keep track of stuff
' ********************************************************************************
' Default menu height
Private Const gMnuHeight                  As Integer = 20
' Menu height
Private Const gMnuWidth                   As Integer = 20
' Menu item captions - stored in an array
Private gMenuCaps                         As Collection
' ID for the last top level menu
Private gLastTopMenuID                    As Long
' ID numbers for the top level menus - For some blasphemous reason I always start with 666
Private Const gStartTopID                 As Long = 666
Private Type RECT
    left                                      As Long
    top                                       As Long
    right                                     As Long
    bottom                                    As Long
End Type
' Get the windows's dimensions using its handle - used when drawing a texture on the
' top level menus.
' Used to get the system's 3D object border width - subtracted from the overall value
Private Const SM_CXBORDER                 As Integer = 5
Private Const SPI_GETNONCLIENTMETRICS     As Integer = 41
' Logical font type used to size a font with CreateFont
Private Type LOGFONT
    lfHeight                                  As Long
    lfWidth                                   As Long
    lfEscapement                              As Long
    lfOrientation                             As Long
    lfWeight                                  As Long
    lfItalic                                  As Byte
    lfUnderline                               As Byte
    lfStrikeOut                               As Byte
    lfCharSet                                 As Byte
    lfOutPrecision                            As Byte
    lfClipPrecision                           As Byte
    lfQuality                                 As Byte
    lfPitchAndFamily                          As Byte
    lfFaceName(1 To 32)                       As Byte
End Type
Private Type NONCLIENTMETRICS
    cbSize                                    As Long
    iBorderWidth                              As Long
    iScrollWidth                              As Long
    iScrollHeight                             As Long
    iCaptionWidth                             As Long
    iCaptionHeight                            As Long
    lfCaptionFont                             As LOGFONT
    iSMCaptionWidth                           As Long
    iSMCaptionHeight                          As Long
    lfSMCaptionFont                           As LOGFONT
    iMenuWidth                                As Long
    iMenuHeight                               As Long
    lfMenuFont                                As LOGFONT
    lfStatusFont                              As LOGFONT
    lfMessageFont                             As LOGFONT
End Type
' Used to get various system parameters and settings
Private Type POINTAPI
    X                                         As Long
    Y                                         As Long
End Type
' Gets the width & height of text in a DC using that DC's currently selected font
' Gets display info
Private Const LOGPIXELSY                  As Integer = 90
' Type that says how big & wide the menu items will be
Private Type MEASUREITEMSTRUCT
    CtlType                                   As Long
    CtlID                                     As Long
    itemID                                    As Long
    itemWidth                                 As Long
    itemHeight                                As Long
    ItemData                                  As Long
End Type
' Structure used when WM_DRAWITEM is passed that says
'  which part of the form/menu/etc will be worked on
Private Type DRAWITEMSTRUCT
    CtlType                                   As Long
    CtlID                                     As Long
    itemID                                    As Long
    itemAction                                As Long
    itemState                                 As Long
    hwndItem                                  As Long
    hdc                                       As Long
    rcItem                                    As RECT
    ItemData                                  As Long
End Type
' Used to copy information from pointers into structures
' ******************************************************************************
' MENU DECLARES - Used to get / set information for menu items
' ******************************************************************************
' Used to set the "Owner drawn" functionality of the menu item
Private Type MENUITEMINFO
    cbSize                                    As Long
    fMask                                     As Long
    fType                                     As Long
    fState                                    As Long
    wid                                       As Long
    hSubMenu                                  As Long
    hbmpChecked                               As Long
    hbmpUnchecked                             As Long
    dwItemData                                As Long
    dwTypeData                                As String
    cch                                       As Long
End Type
Private Const MIIM_STATE                  As Long = &H1
Private Const MIIM_ID                     As Long = &H2
Private Const MIIM_SUBMENU                As Long = &H4
Private Const MIIM_TYPE                   As Long = &H10
Private Const MIIM_DATA                   As Long = &H20
Private Const MF_MENUBREAK                As Long = &H40
Private Const MF_BYPOSITION               As Long = &H400
Private Const MF_OWNERDRAW                As Long = &H100
Private Const MF_SEPARATOR                As Long = &H800
Private Const ODS_SELECTED                As Long = &H1
Private Const ODS_DISABLED                As Long = &H4     ' Whether an item is disabled or not
Private Const ODS_CHECKED                 As Long = &H8     ' Whether an item is "checked" or not
Private Const ODS_HOTTRACK                As Long = &H40    ' Whether the menubar is hottracking

' ******************************************************************************
' SUBCLASSING ROUTINES - All messages are sent directly to the parent form
' ******************************************************************************
Private Const GWL_WNDPROC                 As Long = (-4)
Private Const WM_ERASEBKGND               As Long = &H14
Private Const WM_DRAWITEM                 As Long = &H2B
' Used to actually draw the owner-drawn item
Private Const WM_MEASUREITEM              As Long = &H2C
' Used to return the size of the area in which
' we will be drawing. This would be extremely
' useful if we wanted to create a custom seperator
Private Const WM_INITMENUPOPUP            As Long = &H117
' Used to grab the menus and make them owner-drawn
' ******************************************************************************
' GRAPHICS DECLARES (GDI32 & USER32) - for drawing edge and pictures and text
' ******************************************************************************
' The following are used to draw the upraised square on the menu item:
Private Const BDR_RAISEDINNER             As Long = &H4
Private Const BDR_SUNKENOUTER             As Long = &H2
Private Const BF_BOTTOM                   As Long = &H8
Private Const BF_LEFT                     As Long = &H1
Private Const BF_RIGHT                    As Long = &H4
Private Const BF_TOP                      As Long = &H2
Private Const BF_RECT                     As Double = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Type BITMAP
    bmType                                    As Long
    bmWidth                                   As Long
    bmHeight                                  As Long
    bmWidthBytes                              As Long
    bmPlanes                                  As Integer
    bmBitsPixel                               As Integer
    bmBits                                    As Long
End Type
' *************************
'   For disabled items:
' *************************
Private Const DST_PREFIXTEXT              As Long = &H2
Private Const DST_BITMAP                  As Long = &H4
Private Const DSS_NORMAL                  As Long = &H0
Private Const DSS_DISABLED                As Long = &H20
Private Const ODT_MENU                    As Integer = 1
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, _
                                                                     ByVal W As Long, _
                                                                     ByVal E As Long, _
                                                                     ByVal O As Long, _
                                                                     ByVal W As Long, _
                                                                     ByVal I As Long, _
                                                                     ByVal u As Long, _
                                                                     ByVal s As Long, _
                                                                     ByVal C As Long, _
                                                                     ByVal OP As Long, _
                                                                     ByVal CP As Long, _
                                                                     ByVal Q As Long, _
                                                                     ByVal PAF As Long, _
                                                                     ByVal f As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, _
                                                                      ByVal nPosition As Long, _
                                                                      ByVal wFlags As Long, _
                                                                      ByVal wIDNewItem As Long, _
                                                                      ByVal lpString As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, _
                                                                                ByVal un As Long, _
                                                                                ByVal bool As Boolean, _
                                                                                lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, _
                                                     ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
                                                     ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
                                                                                ByVal un As Long, _
                                                                                ByVal b As Boolean, _
                                                                                lpmii As MENUITEMINFO) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal Msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
                                                qrc As RECT, _
                                                ByVal edge As Long, _
                                                ByVal grfFlags As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nXOrg As Long, _
                                                    ByVal nYOrg As Long, _
                                                    lppt As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String, _
                                                                ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, _
                                                                      ByVal lpString As String) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                    ByVal hBrush As Long, _
                                                                    ByVal lpDrawStateProc As Long, _
                                                                    ByVal lParam As Long, _
                                                                    ByVal wParam As Long, _
                                                                    ByVal X As Long, _
                                                                    ByVal Y As Long, _
                                                                    ByVal cx As Long, _
                                                                    ByVal cy As Long, _
                                                                    ByVal flags As Long) As Long
Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                        ByVal hBrush As Long, _
                                                                        ByVal lpDrawStateProc As Long, _
                                                                        ByVal lString As String, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal X As Long, _
                                                                        ByVal Y As Long, _
                                                                        ByVal cx As Long, _
                                                                        ByVal cy As Long, _
                                                                        ByVal flags As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal X1 As Long, _
                                                ByVal Y1 As Long, _
                                                ByVal X2 As Long, _
                                                ByVal Y2 As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                                                          ByVal hWnd2 As Long, _
                                                                          ByVal lpsz1 As String, _
                                                                          ByVal lpsz2 As String) As Long
Private Function drawMnuBitmap(ByVal lngHwnd As Long, _
                               ByVal drawInfoPtr As Long) As Long

Dim aDIS           As DRAWITEMSTRUCT
Dim lPic           As StdPicture   ' our picture
Dim lMask          As StdPicture   ' our picture's mask
Dim picDC          As Long         ' picture device context
Dim maskDC         As Long         ' mask device context
Dim sCap           As String       ' Caption string
Dim lPicInds       As Variant      ' picture indices (array where 0=picID, 1=maskID)
Dim boxRect        As RECT         ' a rectangle used to paint stuff
Dim abrush         As Long         ' brush object
Dim noPics         As Boolean      ' true if no pictures for this menu item
Dim colDC          As Long         ' a color's DC
Dim colBmp         As Long         ' a Colors bitmap - used for adjusting stuff
Dim aRect          As RECT         ' another rectangle. imagine that.
Dim sAcc           As String       ' accelerator string
Dim aPt            As POINTAPI     ' a user defined structure for x,y
Dim aPen           As Long         ' a pen object
Dim customFont     As Long         ' custom font (if specified)
Dim origFont       As Long         ' original font (if customFont is used)
Dim aBmp           As BITMAP
Dim mnuItemWid     As Long
Dim isSep          As Boolean
Static lastDISRect As RECT
Static lastDISID   As Long
Dim backBuffDC     As Long
Dim backBuffBmp    As Long
Dim aBrushOrg      As POINTAPI
Dim backBuffRECT   As RECT
Dim aTempBrush     As Long
Dim MII            As MENUITEMINFO
Dim aStr           As String
Dim lineRECT       As RECT
Dim lineBrush      As Long
Dim nval           As Long
    If drawInfoPtr = 0 Then
        Exit Function
    End If
    ' get the drawing structure information
    CopyMemory aDIS, ByVal drawInfoPtr, LenB(aDIS)
    aStr = String$(128, 0)
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_TYPE Or MIIM_ID
        .cch = 127
        .dwTypeData = aStr
    End With
    GetMenuItemInfo aDIS.hwndItem, aDIS.itemID, False, MII
    isSep = ((MII.fType And MF_SEPARATOR) = MF_SEPARATOR)
    ' Create a back buffer to draw stuff on - prevents flickering
    backBuffRECT.right = aDIS.rcItem.right - aDIS.rcItem.left
    backBuffRECT.bottom = aDIS.rcItem.bottom - aDIS.rcItem.top
    backBuffDC = CreateCompatibleDC(aDIS.hdc)
    backBuffBmp = CreateCompatibleBitmap(aDIS.hdc, backBuffRECT.right, backBuffRECT.bottom)
    DeleteObject SelectObject(backBuffDC, backBuffBmp)
    mnuItemWid = gMnuWidth
    'Checked ?
    abrush = CreatePatternBrush(CustomMenu.Picture.Handle)
'-----------------------------------------------------------------------------------
' If this is a top level menu and its the first one on a new row,
' go back and paint from the end of the previous item all the way to
' the right side. This is different from the very last item.
'-----------------------------------------------------------------------------------
    If aDIS.ItemData = 2 Then

        With lastDISRect
            If (.top < aDIS.rcItem.top And aDIS.itemID > gStartTopID And .right <> 0 And aDIS.itemID = lastDISID + 1) Then
                boxRect.left = .right
                boxRect.top = .top
                GetWindowRect lngHwnd, aRect
                boxRect.right = aRect.right - aRect.left - GetSystemMetrics(SM_CXBORDER) * 2 - 2
                boxRect.bottom = .bottom
                FillRect aDIS.hdc, boxRect, abrush
            End If
        End With 'lastDISRect
        CopyMemory lastDISRect, aDIS.rcItem, Len(lastDISRect)
        lastDISID = aDIS.itemID
    End If
'-------------------------------------------------------------------------
' If this is a top level menu and its the last one, then we should
' temporarily change the item width so it paints all the way to the right
' on the form
'-------------------------------------------------------------------------
    If (aDIS.ItemData = 2 And aDIS.itemID = gLastTopMenuID) Then
        GetWindowRect lngHwnd, aRect
        With boxRect
            .left = aDIS.rcItem.right
            .right = aRect.right - aRect.left - GetSystemMetrics(SM_CXBORDER) * 2 - 2
            .top = aDIS.rcItem.top
            .bottom = aDIS.rcItem.bottom
        End With
        FillRect aDIS.hdc, boxRect, abrush
    End If
    ' Adjust our brush's point of origin so it paints correctly
    SetBrushOrgEx backBuffDC, -aDIS.rcItem.left, -aDIS.rcItem.top, aBrushOrg
    aTempBrush = SelectObject(backBuffDC, abrush)
    FillRect backBuffDC, backBuffRECT, abrush
    SelectObject backBuffDC, aTempBrush
    DeleteObject abrush
' get the caption of the menu item we need to draw.
' We will need this not only to draw in the menu, but
' to retrieve the resource IDs for the bitmap images
'===========================================================================
'SEPARATOR
'===========================================================================
    If isSep Then
        With boxRect
            .left = backBuffRECT.left + CustomMenu.PosX '28
            .top = (backBuffRECT.bottom - 2) / 2
            .right = backBuffRECT.right - 5
            .bottom = backBuffRECT.bottom - 2
        End With
        DrawEdge backBuffDC, boxRect, BDR_RAISEDINNER Or BDR_SUNKENOUTER, BF_TOP
        GoTo drawMnuBitmap_exitFunction
    Else
        On Error Resume Next
        sCap = gMenuCaps("M" & CStr(aDIS.itemID) & "-" & lngHwnd)
        If Err.Number <> 0 Then
            sCap = ""
        End If
'==========================================================
' KEY ACCELERATOR \ SHORTCUT MENU
'==========================================================
        If InStr(sCap, vbTab) > 0 Then
            sAcc = VBA.right(sCap, Len(sCap) - InStr(sCap, vbTab))
            sCap = VBA.left(sCap, InStr(sCap, vbTab) - 1)
        End If
    End If
    customFont = getFontForItem(sCap, backBuffDC)
    If customFont <> 0 Then
        origFont = SelectObject(backBuffDC, customFont)
    End If
    If aDIS.ItemData = 2 Then
        GoTo drawMnuBitmap_MenuBarMenu
    End If
    ' Get the bitmap image indices if they exist.
    ' if not, then skip the picture drawing stuff
    On Error Resume Next
    lPicInds = CustomMenu.Icon(sCap)
    If Err.Number <> 0 And (aDIS.itemState And ODS_CHECKED) <> ODS_CHECKED Then
        noPics = True
        Err.Clear
        GoTo drawMnuBitmap_picsDone
    End If
    ' Get the pictures from the resource file
    If (aDIS.itemState And ODS_CHECKED) = ODS_CHECKED Then
        Set lPic = LoadResPicture("checked", vbResBitmap)
        Set lMask = LoadResPicture("checked", vbResBitmap)
        GetObject lPic, Len(aBmp), aBmp
        ' The following is purely to get the checkmark the color it
        colDC = CreateCompatibleDC(aDIS.hdc)
        colBmp = CreateCompatibleBitmap(aDIS.hdc, aBmp.bmWidth, aBmp.bmHeight)
        DeleteObject SelectObject(colDC, colBmp)
        aRect.right = aBmp.bmWidth
        aRect.bottom = aBmp.bmHeight
        abrush = CreateSolidBrush(IIf((aDIS.itemState And ODS_SELECTED) = ODS_SELECTED, CustomColor.DefTextColor, IIf(CustomMenu.UseCustomFonts, CustomColor.ForeColor, CustomColor.DefTextColor)))
        FillRect colDC, aRect, abrush
        DeleteObject abrush
    Else
        Set lPic = LoadResPicture(lPicInds(0), vbResBitmap)
        Set lMask = LoadResPicture(lPicInds(1), vbResBitmap)
        GetObject lPic, Len(aBmp), aBmp
    End If
    ' Create a compatible device context for both of the bitmaps
    With aDIS
        picDC = CreateCompatibleDC(.hdc)
        maskDC = CreateCompatibleDC(.hdc)
        ' select the bitmaps into our device context and delete the temporary 1x1
        ' that's created with the DC
        If (.itemState And ODS_DISABLED) <> ODS_DISABLED Then
            DeleteObject SelectObject(picDC, lPic.Handle)
            DeleteObject SelectObject(maskDC, lMask.Handle)
        End If
    End With 'aDIS
    If (aDIS.itemState And ODS_CHECKED) = ODS_CHECKED Then
        'Make the checkmark the right color
        BitBlt picDC, 0, 0, aBmp.bmWidth, aBmp.bmHeight, colDC, 0, 0, vbSrcPaint
        DeleteDC colDC
        DeleteObject colBmp
        'DrawEdge backBuffDC, backBuffRECT, BDR_SUNKENOUTER, BF_RECT
    End If
'==============================================
' 3D BORDER 'set up a rectangle to draw the upraised edge
'==============================================
    With boxRect
        .top = 0
        .left = CustomMenu.PosX
        .right = .left + mnuItemWid
        .bottom = backBuffRECT.bottom 'boxRect.top + gMnuHeight
    End With
    If (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED And (aDIS.itemState And ODS_CHECKED) <> ODS_CHECKED And (aDIS.itemState And ODS_DISABLED) <> ODS_DISABLED Then
        DrawEdge backBuffDC, boxRect, BDR_RAISEDINNER, BF_RECT
    End If
drawMnuBitmap_picsDone:
' If the item is in a "highlighted" state, then
    If (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED And Not (aDIS.itemState And ODS_DISABLED) = ODS_DISABLED Then
' Draw the edge
'=================================================
' HILIGHT COLOR 'set up a rectangle to draw the "highlight" color
'=================================================
        If MenuMode = XPlook Then
           nval = 1
        Else
           nval = 0
        End If
        With boxRect
            If MenuMode = XPlook Then
            .left = 25 ' IIf(noPics, backBuffRECT.left + CustomMenu.PosX, .right) + 1
            Else
             .left = IIf(noPics, backBuffRECT.left + CustomMenu.PosX, .right) + 1
            End If
            .top = backBuffRECT.top + nval
            .bottom = backBuffRECT.bottom - nval
            .right = backBuffRECT.right - nval
        ' create a brush in the highlight color
        End With
        With lineRECT
             .left = boxRect.left - 1 '(noPics, backBuffRECT.left + CustomMenu.PosX, .right) + 1
             .top = boxRect.top - 1 ' backBuffRECT.top
             .bottom = boxRect.bottom + 1 ' backBuffRECT.bottom
             .right = boxRect.right + 1 '.right
        End With
        If MenuMode = XPlook Then
            lineBrush = CreateSolidBrush(CustomColor.RECTColor)
            FillRect backBuffDC, lineRECT, lineBrush
        End If
        abrush = CreateSolidBrush(CustomColor.HilightColor)
        
        ' color the rectangular area
        FillRect backBuffDC, boxRect, abrush
        ' delete the brush object (clear up resources)
        DeleteObject abrush
    End If
    'Set our text colors appropriately, depending on whether we are
    ' in a highlighted state or not
    If (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED Then
        SetTextColor backBuffDC, CustomColor.SelectedTextColor
        SetBkColor backBuffDC, CustomColor.HilightColor
    Else
        SetTextColor backBuffDC, IIf(CustomColor.ForeColor = 0, CustomColor.DefTextColor, CustomColor.ForeColor)
        SetBkColor backBuffDC, CustomColor.BackColor
    End If
'====================================
'Print the text \Print the accelerator (Ctl-Whatever)
'====================================
    SetBkMode backBuffDC, 0
    GetTextExtentPoint32 backBuffDC, sCap, Len(sCap), aPt
    With boxRect
        .top = backBuffRECT.top + ((backBuffRECT.bottom - backBuffRECT.top - aPt.Y) / 2)
        .bottom = backBuffRECT.bottom
        .right = backBuffRECT.right - 5
        .left = mnuItemWid + CustomMenu.PosX + 4
        DrawStateText backBuffDC, CustomMenu.PosX + 4, 0, sCap, Len(sCap), .left, .top, 0, 0, DST_PREFIXTEXT Or IIf((aDIS.itemState And ODS_DISABLED) = ODS_DISABLED, DSS_DISABLED, DSS_NORMAL)
    End With
    If LenB(sAcc) Then
        DeleteObject SelectObject(backBuffDC, origFont)
        customFont = getFontForItem(sCap, backBuffDC, True)
        origFont = SelectObject(backBuffDC, customFont)
        GetTextExtentPoint32 backBuffDC, sAcc, Len(sAcc), aPt
        DrawStateText backBuffDC, 0, 0, sAcc, Len(sAcc), boxRect.right - aPt.X, boxRect.top, 0, 0, DST_PREFIXTEXT Or IIf((aDIS.itemState And ODS_DISABLED) = ODS_DISABLED, DSS_DISABLED, DSS_NORMAL)
    End If
'============================================================================================================================================
    If lPic Is Nothing Then
        GoTo drawMnuBitmap_exitFunction
    End If
    If lPic.Handle <> 0 Then
        If (aDIS.itemState And ODS_DISABLED) = ODS_DISABLED Then
            DrawState backBuffDC, 0, 0, lPic.Handle, 0, CustomMenu.PosX + 2, backBuffRECT.top + (backBuffRECT.bottom - backBuffRECT.top - aBmp.bmHeight) / 2, aBmp.bmHeight, aBmp.bmHeight, DST_BITMAP Or DSS_DISABLED
        Else
            ' Blt the mask
            BitBlt backBuffDC, CustomMenu.PosX + 2, backBuffRECT.top + (backBuffRECT.bottom - backBuffRECT.top - aBmp.bmHeight) / 2, aBmp.bmWidth, aBmp.bmHeight, maskDC, 0, 0, vbMergePaint
            ' Blt the picture
           BitBlt backBuffDC, CustomMenu.PosX + 2, backBuffRECT.top + (backBuffRECT.bottom - backBuffRECT.top - aBmp.bmHeight) / 2, aBmp.bmWidth, aBmp.bmHeight, picDC, 0, 0, vbSrcAnd
        End If
        ' Clean up our graphics resources.
        DeleteDC picDC
        DeleteObject lPic.Handle
        DeleteDC maskDC
        DeleteObject lMask.Handle
    End If
drawMnuBitmap_exitFunction:
' Check to see if we are in a 2nd or 3rd column... if so, then
' we need to draw over a little bit of the menu.
    With aDIS
        If .ItemData <> 2 Then
            If .rcItem.left > 0 Then
                ' Calculate the rectangle for 4 pixels to the left
                aRect.left = .rcItem.left - 4
                aRect.right = .rcItem.left
                aRect.top = .rcItem.top
                aRect.bottom = .rcItem.bottom
                DeleteObject abrush
                ' Create a new brush object
                If Not (CustomMenu.Picture Is Nothing) Then
                    'Pattern brush
                    abrush = CreatePatternBrush(CustomMenu.Picture.Handle)
                End If
                ' Fill in the area
                FillRect .hdc, aRect, abrush
                ' Clean up our brush resource.
                DeleteObject abrush
            End If
        End If
    End With 'aDIS
    ' Clean up our graphics resources to free up memory
    If origFont <> 0 Then
        SelectObject backBuffDC, origFont
        DeleteObject customFont
    End If
    DeleteObject abrush
    DeleteDC picDC
    DeleteDC maskDC
    ' Transfer the menu item from our back buffer into the menu DC
    BitBlt aDIS.hdc, aDIS.rcItem.left, aDIS.rcItem.top, backBuffRECT.right, backBuffRECT.bottom, backBuffDC, 0, 0, vbSrcCopy
    DeleteDC backBuffDC
    DeleteObject backBuffBmp
    On Error GoTo 0
Exit Function
drawMnuBitmap_MenuBarMenu:
    ' Top level menus... These things are so freakin easy its funny.
    If (aDIS.itemState And ODS_HOTTRACK) = ODS_HOTTRACK Then
    ' This little style bit is courtesy of VolteFace from www.visualbasicforum.com
        DrawEdge backBuffDC, backBuffRECT, BDR_RAISEDINNER, BF_RECT
    ElseIf (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED And (aDIS.itemState And ODS_DISABLED) <> ODS_DISABLED Then
    ' If its a selected item, paint the background with the systems 'Highlighted' color

        With CustomColor
            abrush = SelectObject(backBuffDC, CreateSolidBrush(.HilightColor))
            ' Also make the text print out in the highlighted text color
            SetTextColor backBuffDC, .SelectedTextColor
            aPen = SelectObject(backBuffDC, CreatePen(0, 1, .BorderColor))
        End With 'CustomColor
        Rectangle backBuffDC, backBuffRECT.left, backBuffRECT.top, backBuffRECT.right, backBuffRECT.bottom
        DeleteObject SelectObject(backBuffDC, aPen)
        DeleteObject SelectObject(backBuffDC, abrush)
    Else
        ' otherwise just make it the system menu text color
        'SetTextColor aDIS.hdc, GetSysColor(IIf((aDIS.itemState And ODS_DISABLED) = ODS_DISABLED, 17, 7))
        SetTextColor backBuffDC, IIf(CustomColor.ForeColor = 0, CustomColor.MenuTextColor, CustomColor.ForeColor)
    End If
    ' Make the text print transparently
    SetBkMode backBuffDC, 0
    ' Get text dimensions
    GetTextExtentPoint32 backBuffDC, sCap, Len(sCap), aPt
    ' Draw the text!
    DrawStateText backBuffDC, 0, 0, sCap, Len(sCap), backBuffRECT.top + (backBuffRECT.right - backBuffRECT.left - aPt.X) / 2, backBuffRECT.top + (backBuffRECT.bottom - backBuffRECT.top - aPt.Y) / 2, 0, 0, DST_PREFIXTEXT Or IIf((aDIS.itemState And ODS_DISABLED) = ODS_DISABLED, DSS_DISABLED, DSS_NORMAL)
    GoTo drawMnuBitmap_exitFunction
End Function
Private Function getFontForItem(ByVal anID As String, _
                                lngHdc As Long, _
                                Optional ByVal getDefault As Boolean) As Long

Dim aNCMETRIC      As NONCLIENTMETRICS
Dim logPixConv     As Double
Dim systemMenuFont As String


    ' Calculate a logical pixels conversion factor
    logPixConv = GetDeviceCaps(lngHdc, LOGPIXELSY) / 72
    On Error Resume Next
    ' Get the nonclient metrics, including system menu height
    aNCMETRIC.cbSize = Len(aNCMETRIC)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, aNCMETRIC.cbSize, aNCMETRIC, 0

    ' If we're not using custom fonts, or they are not set just use the normal fonts.
    If Not CustomMenu.UseCustomFonts Then
      
        'DEFAULT FONT \ SYSTEM MENUFONT
        With aNCMETRIC.lfMenuFont
            systemMenuFont = StrConv(.lfFaceName, vbUnicode)
            systemMenuFont = left(systemMenuFont, InStr(systemMenuFont, Chr$(0)) - 1)
            getFontForItem = CreateFont(-1 * .lfHeight * logPixConv, .lfWidth, .lfEscapement, .lfOrientation, .lfWeight, .lfItalic, .lfUnderline, .lfStrikeOut, .lfCharSet, .lfOutPrecision, .lfClipPrecision, .lfQuality, .lfPitchAndFamily, systemMenuFont)
        End With
    Else
        'FONT CUSTIMIZE
        With aNCMETRIC.lfMenuFont
            If CustomMenu.FontBold Then
               getFontForItem = CreateFont(-1 * .lfHeight * logPixConv, .lfWidth, .lfEscapement, .lfOrientation, .lfWeight * 2, CustomMenu.FontItalic, CustomMenu.FontUnderline, CustomMenu.FontStrikeOut, .lfCharSet, .lfOutPrecision, .lfClipPrecision, .lfQuality, .lfPitchAndFamily, CustomMenu.FontName)
            Else
               getFontForItem = CreateFont(-1 * .lfHeight * logPixConv, .lfWidth, .lfEscapement, .lfOrientation, .lfWeight, CustomMenu.FontItalic, CustomMenu.FontUnderline, CustomMenu.FontStrikeOut, .lfCharSet, .lfOutPrecision, .lfClipPrecision, .lfQuality, .lfPitchAndFamily, CustomMenu.FontName)
            End If
        End With
    End If
    On Error GoTo 0
End Function
Public Function getMenuDimensions(lngHwnd As Long, _
                                  ByVal subItemID As Long, _
                                  ByVal itemType As Long, _
                                  aPt As POINTAPI) As Boolean
Dim aPt2     As POINTAPI
Dim aCap     As String
Dim formDC   As Long
Dim origFont As Long
Dim mnuFont  As Long
Dim lPic     As StdPicture
Dim aBmp     As BITMAP
Dim lPicInds As Variant
Dim aWid     As Long
' Determine an estimate for the menu width based on the stored menu
'  caption string.
    On Error Resume Next
    aCap = gMenuCaps("M" & CStr(subItemID) & "-" & lngHwnd)
    ' Replace tab character with a space.
    If InStr(aCap, vbTab) > 0 Then
        aCap = VBA.left(aCap, InStr(aCap, vbTab) - 1) & " " & VBA.right(aCap, Len(aCap) - InStr(aCap, vbTab))
    End If
    ' Get the form's DC using its handle
    formDC = GetDC(lngHwnd)
    ' Get the proper font size
    mnuFont = getFontForItem(aCap, formDC)
    ' select the font into the device context
    origFont = SelectObject(formDC, mnuFont)
    ' Get the text width using the form's DC as a reference
    GetTextExtentPoint32 formDC, aCap, Len(aCap), aPt2
    ' Replace the font with the original
    SelectObject formDC, origFont
    ' Release the form's DC back to itself
    ReleaseDC lngHwnd, formDC
    ' Delete the temporary font
    DeleteObject mnuFont
    Err.Clear
    ' Check and see if the bitmap is taller than the font
    aCap = gMenuCaps(CStr(subItemID))
    If InStr(aCap, vbTab) > 0 Then
        aCap = VBA.left(aCap, InStr(aCap, vbTab) - 1)
    End If
    lPicInds = CustomMenu.Icon(aCap)
    aWid = gMnuWidth + 30
    If Err.Number = 0 Then
        Set lPic = LoadResPicture(lPicInds(0), vbResBitmap)
        GetObject lPic.Handle, Len(aBmp), aBmp
        DeleteObject lPic.Handle
    End If
    '  Make the width = text width plus 2 times the menu height (and picture width ;) )
    aPt.X = (aPt2.X + IIf(itemType = 2, 0, aWid) + 6)
    '  Calculate the height
    aPt.Y = IIf(aBmp.bmHeight > aPt2.Y, aBmp.bmHeight, aPt2.Y) + 6
    On Error GoTo 0
End Function

Private Function ODWindowProc(ByVal lngHwnd As Long, _
                              ByVal uMsg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long

Dim aWid       As Long
Dim hSysMenu   As Long
Dim aForm      As Form
Dim isSysMenu  As Boolean
Dim aMeas      As MEASUREITEMSTRUCT
Dim aMnuDim    As POINTAPI
Dim oldWndProc As Long
Dim aDIS       As DRAWITEMSTRUCT
Dim abrush     As Long
Dim aRect      As RECT
Dim aHWNDTmp   As Long
Dim sClassName As String
Dim isSep      As Boolean
    oldWndProc = GetProp(lngHwnd, "ODMenuOrigProc")
    If uMsg = WM_DRAWITEM Then
        CopyMemory aDIS, ByVal lParam, Len(aDIS)
        If aDIS.CtlType = ODT_MENU Then
        'Use our custom drawing subroutine to draw the menu item
            drawMnuBitmap lngHwnd, lParam
            'Don't do any other processing
            ODWindowProc = False
        Else
            ODWindowProc = CallWindowProc(oldWndProc, lngHwnd, uMsg, wParam, lParam)
        End If
    ElseIf uMsg = WM_ERASEBKGND Then
    ' Figure out if this is a toolbar.
        aHWNDTmp = WindowFromDC(wParam)
        sClassName = String$(128, Chr$(0))
        GetClassName aHWNDTmp, sClassName, Len(sClassName)
        ' If it IS a toolbar and we have a custom menu background selected
        If InStr(sClassName, "toolbar") > 1 Then
            GetWindowRect aHWNDTmp, aRect
            With aRect
                .right = .right - .left
                .bottom = .bottom - .top
                .top = 0
                .left = 0
            End With 'aRect
            If Not (CustomMenu.Picture Is Nothing) Then
                abrush = CreatePatternBrush(CustomMenu.Picture.Handle)
            Else
                abrush = CreateSolidBrush(CustomColor.BackColor)
            End If
            FillRect wParam, aRect, abrush
            DeleteObject abrush
            ODWindowProc = True
        Else
            ODWindowProc = CallWindowProc(oldWndProc, lngHwnd, uMsg, wParam, lParam)
        End If
    ElseIf uMsg = WM_MEASUREITEM Then
    'copy the information from pointer lparam to our structure
        CopyMemory aMeas, ByVal lParam, Len(aMeas)
        If aMeas.CtlType = ODT_MENU Then
        'determine the width of our menu item
            getMenuDimensions lngHwnd, aMeas.itemID, aMeas.ItemData, aMnuDim
            aWid = aMnuDim.X
            'if this item value is bigger than the previous one, store the bigger
            'value. This allows the menu to be properly sized for the biggest item.
            'If aMeas.itemWidth < aWid Then aMeas.itemWidth = aWid
            aMeas.itemWidth = aWid
            '  Make each item either 6 or gMnuHeight pixels high. We check the
            '  item data value to determine whether it is a seperator or a
            '  regular menu (0 = seperator, 1 = normal menu)
            aMeas.itemHeight = IIf(isSep, 6, aMnuDim.Y)
            'Copy the structure back to the one located at pointer location
            CopyMemory ByVal lParam, aMeas, Len(aMeas)
            'Don't do any other processing
            ODWindowProc = False
        Else
            ODWindowProc = CallWindowProc(oldWndProc, lngHwnd, uMsg, wParam, lParam)
        End If
    ElseIf uMsg = WM_INITMENUPOPUP Then
    ' Make sure that we're not trying to set OD styles on the system menu...
    '  (you REALLY REALLY REALLY don't want to try to do that)
        isSysMenu = False
        For Each aForm In Forms
            hSysMenu = GetSystemMenu(aForm.hwnd, False)
            ' if its not the systemmenu then set all the styles to ownerdrawn and
            '  pop it open!
            isSysMenu = isSysMenu Or (wParam = hSysMenu)
        Next aForm
        ' Invoke whatever it was going to do
        ODWindowProc = CallWindowProc(oldWndProc, lngHwnd, uMsg, wParam, lParam)
        If Not isSysMenu Then
            setPopupStyleOD wParam, lngHwnd
        End If
    Else
        ' Invoke the default window procedure
        ODWindowProc = CallWindowProc(oldWndProc, lngHwnd, uMsg, wParam, lParam)
    End If
End Function
Private Sub setPopupStyleOD(aHwnd As Long, _
                            ByVal wndHwnd As Long, _
                            Optional ByVal anItemInd As Long)
Dim I         As Long
Dim MII       As MENUITEMINFO
Dim capStr    As String
Dim startInd  As Long
Dim endInd    As Long
Dim lNewStyle As Long
' Determine whether we are going set the ownerdrawn for one individual item
' or for a whole submenu.
    If anItemInd > 0 Then
        startInd = anItemInd
        endInd = anItemInd
    Else
        startInd = 0
        endInd = GetMenuItemCount(aHwnd) - 1
    End If
    ' Loop thru from startInd to endInd
    For I = startInd To endInd
    ' initialize our menu item info structure to get data
        With MII
            ' MII.fMask = MIIM_DATA Or MIIM_ID Or MIIM_SUBMENU Or MIIM_STATE Or MIIM_TYPE
            .fMask = MIIM_TYPE Or MIIM_ID
            .cch = 127
            .dwTypeData = String$(128, 0)
            .cbSize = Len(MII)
            ' get the menu item information
        End With
        GetMenuItemInfo aHwnd, I, True, MII
        ' get the ID number for the menu item
        GetMenuItemID aHwnd, I
        ' determine if the item is a seperator or not
        With MII
            'isSep = ((.fType And MF_SEPARATOR) = MF_SEPARATOR)
            lNewStyle = .fType Or MF_OWNERDRAW Or 0& 'Or mii.fState
            ' trim extra null characters out of the caption
            If InStr(.dwTypeData, Chr$(0)) > 0 Then
                capStr = left(.dwTypeData, InStr(.dwTypeData, Chr$(0)) - 1)
                ' Split menu item if first char is a pipe
                If left(capStr, 1) = "|" Then
                    capStr = right(capStr, Len(capStr) - 1)
                    lNewStyle = lNewStyle Or MF_MENUBREAK
                End If
            Else
                capStr = ""
            End If
        End With
        MII.fType = lNewStyle
        MII.fMask = MIIM_TYPE Or MIIM_ID 'Or MIIM_DATA
        SetMenuItemInfo aHwnd, I, True, MII
        ' provided there is a caption, store it. This is a weird requirement
        '  I ran into while testing out the OD menus in NT. For some reason,
        '  MII.dwTypeData doesn't always have the caption string... sometimes
        '  it disappears! That's why we can't just use the information in the
        '  DRAWITEMSTRUCT in our ownerdrawn procedure.
        If LenB(capStr) Then
            On Error Resume Next
        ' Store an item unique to each WINDOW AND ITEM ID
            gMenuCaps.Remove "M" & CStr(MII.wid) & "-" & wndHwnd
            gMenuCaps.Add capStr, "M" & CStr(MII.wid) & "-" & wndHwnd
            On Error GoTo 0
        End If
    Next I
End Sub
Public Sub startODMenus(aControl As Object, ByVal bMenuBar As Boolean)

Dim hwnd     As Long
Dim origProc As Long

    If CustomMenu.Icon Is Nothing Then
        Set CustomMenu.Icon = New Collection
    End If
    If gMenuCaps Is Nothing Then
        Set gMenuCaps = New Collection
    End If
    If CustomMenu.CustomFonts Is Nothing Then
        Set CustomMenu.CustomFonts = New Collection
    End If
If TypeOf aControl Is Form Then
    ' If the user specifies to make the menubar ownerdrawn, do it!
        If bMenuBar Then
          'makeTopMenusOD aControl
          ' burasý
          MenuArkaPlan aControl
        End If
        hwnd = aControl.hwnd
    End If
    ' Start the subclassing
    origProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf ODWindowProc)
    ' Store the original process address in Windows' catalog using the form's handle
    SetProp hwnd, "ODMenuOrigProc", origProc
End Sub
Public Sub stopODMenus(aControl As Object)
Dim origProc As Long
Dim hwnd     As Long
    If TypeOf aControl Is Toolbar Then
        hwnd = FindWindowEx(aControl.hwnd, ByVal 0&, "msvb_lib_toolbar", vbNullString)
    ElseIf TypeOf aControl Is Form Then
        ' Get the original process address using the form's handle
        If Forms.Count = 1 Then
            'last form - unhook stuff and clear out the crap
            Set gMenuCaps = Nothing
            Set CustomMenu.Icon = Nothing
            Set CustomMenu.CustomFonts = Nothing
        End If
        hwnd = aControl.hwnd
    End If
    origProc = GetProp(hwnd, "ODMenuOrigProc")
    ' Unsubclass the form by replacing the original process address
    SetWindowLong hwnd, GWL_WNDPROC, origProc
    ' Remove the property entry from the Windows' catalog
    RemoveProp hwnd, "ODMenuOrigProc"
End Sub

Public Sub MenuArkaPlan(HangiForm As Form)

Dim MI As MENUINFO
With MI
.cbSize = Len(MI)
.fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
.hbrBack = CreateSolidBrush(&H404040)
SetMenuInfo GetMenu(HangiForm.hwnd), MI

End With
End Sub

Public Sub makeTopMenusOD(aForm As Form)
Dim hMenubar   As Long
Dim numTopMnus As Long
Dim aMII       As MENUITEMINFO
Dim I          As Long
Dim sCap       As String
Dim aStart     As Long
    ' Grab the form's menubar
    hMenubar = GetMenu(aForm.hwnd)
    ' Get the number of top-level menubar items
    numTopMnus = GetMenuItemCount(hMenubar)
    ' store the last ID number just for quick reference in the drawing routine
    gLastTopMenuID = gStartTopID + numTopMnus - 1
    aStart = IIf(aForm.WindowState = vbMaximized, 1, 0)
    For I = aStart To numTopMnus - 1
    ' initialize our menu item info structure to get data
        With aMII
            .fMask = MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_TYPE
            .cch = 127
            .dwTypeData = String$(128, 0)
            .cbSize = Len(aMII)
        ' Actually go get the menu item data
        End With 'aMII
        GetMenuItemInfo hMenubar, I, True, aMII
        ' Save the captions in our memory collection
        On Error Resume Next
        sCap = VBA.left(aMII.dwTypeData, aMII.cch)
        If LenB(sCap) Then
            gMenuCaps.Remove "M" & CStr(gStartTopID + I) & "-" & aForm.hwnd
            gMenuCaps.Add sCap, "M" & CStr(gStartTopID + I) & "-" & aForm.hwnd
        End If
        On Error GoTo 0
        'Get the state of the menu item
        aMII.fMask = MIIM_STATE
        GetMenuItemInfo hMenubar, I, True, aMII
        ' Turn the menubar item into an owner-drawn one
         ModifyMenu hMenubar, I, MF_OWNERDRAW Or MF_BYPOSITION Or aMII.fState, gStartTopID + I, ByVal 2&
        SetMenuItemInfo hMenubar, I, True, aMII
    Next I
End Sub

