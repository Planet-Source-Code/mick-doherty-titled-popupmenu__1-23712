Attribute VB_Name = "PopMenu"
Declare Function CreatePopupMenu Lib "user32" () As Long

Declare Function TrackPopupMenu Lib "user32" _
        (ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nReserved As Long, _
        ByVal hwnd As Long, _
        ByVal lprc As Any) As Long
        
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
        (ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpNewItem As Any) As Long
        
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
        (ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpString As Any) As Long
        
Public Declare Function DestroyMenu Lib "user32" _
        (ByVal hMenu As Long) As Long
        
Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function BitBlt Lib "gdi32" _
        (ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Declare Function SetRect Lib "user32" (lpRect As RECT, _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long) As Long

Declare Function DrawCaption Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal hdc As Long, _
        pcRect As RECT, _
        ByVal un As Long) As Long
        
Declare Function GetMenuItemRect Lib "user32" _
        (ByVal hwnd As Long, ByVal hMenu As Long, _
        ByVal uItem As Long, _
        lprcItem As RECT) As Long

Declare Function GetMenuItemCount Lib "user32" _
        (ByVal hMenu As Long) As Long

Declare Function GetPixel Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long) As Long
        
Declare Function SetPixel Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal crColor As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type

Const MF_APPEND = &H100&
Const MF_BYCOMMAND = &H0&
Const MF_BYPOSITION = &H400&
Const MF_DEFAULT = &H1000&
Const MF_DISABLED = &H2&
Const MF_ENABLED = &H0&
Const MF_GRAYED = &H1&
Const MF_MENUBARBREAK = &H20&
Const MF_OWNERDRAW = &H100&
Const MF_POPUP = &H10&
Const MF_REMOVE = &H1000&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Const MF_UNCHECKED = &H0&
Const MF_BITMAP = &H4&
Const MF_USECHECKBITMAPS = &H200&

Public Const MF_CHECKED = &H8&
Public Const MFT_RADIOCHECK = &H200&

Const TPM_RETURNCMD = &H100&

Const DC_GRADIENT = &H20
Const DC_ACTIVE = &H1
Const DC_ICON = &H4
Const DC_SMALLCAP = &H2
Const DC_TEXT = &H8

Public hMenu As Long
Public hSubmenu As Long
Public chkMnuFlags(2) As Long
Public MP As PointAPI, sMenu As Long
Public mnuHeight As Single

Public Sub MeasureMenu(ByRef lP As Long)
    
    'It would appear that you cannot actually get measurements here,
    'you can only set them. There are no measurements until after the
    'Menu is drawn, but you only get a WM_MEASUREITEM message before the
    'initial WM_DRAWITEM.
    
    Dim MIS As MEASUREITEMSTRUCT
    'Load MIS with that in memory
    CopyMemory MIS, ByVal lP, Len(MIS)
        MIS.itemWidth = 5   '(18 - 1) - 12. I don't know where the 12 comes
        'from, but there always seems to be 12 pixels more than I want.
        '18 is Small Titlebar height.
    
    'Return the updated MIS
    CopyMemory ByVal lP, MIS, Len(MIS)
    
End Sub

Public Sub DrawMenu(ByRef lP As Long)
    
    Dim DIS As DRAWITEMSTRUCT, rct As RECT, lRslt As Long
    
    CopyMemory DIS, ByVal lP, Len(DIS)
    
    With AppForm
        'since we can't measure in the MeasureMenu sub we'll do it here.
        'we cannot just get the sidebar height as it will only return
        'the height of an empty menu item. (i.e. 13). Maybe we can get the
        'height of the whole menu with some other API call that I don't know
        'about. I tried GetWindowRect.
        
        'String Menus
        GetMenuItemRect .hwnd, hMenu, 1, rct
        mnuHeight = (rct.Bottom - rct.Top) * (GetMenuItemCount(hMenu) - GetMenuItemCount(hSubmenu) - 1)
        'Separators
        GetMenuItemRect .hwnd, hMenu, 3, rct
        mnuHeight = mnuHeight + (rct.Bottom - rct.Top) * 2 '2 Seperators
        
        'set the size of our sidebar
        SetRect rct, 0, 0, mnuHeight, 18
        
        'This is a bit of a copout, but it works
        'You could always use GradientFillRect and then draw rotated text
        'straight onto the sidebar, but this is much easier
        'you could use a hidden picturebox for this
        'Draw a form caption onto our userform, the length of our menu height
        DrawCaption .hwnd, .hdc, rct, DC_SMALLCAP Or DC_ACTIVE Or DC_TEXT Or DC_GRADIENT
        
        Dim X As Single, Y As Single
        Dim nColor As Long
        
        'rotate our caption through 270 degrees
        'and paint onto menu
        For X = 0 To mnuHeight
            For Y = 0 To 17
                nColor = GetPixel(.hdc, X, Y)
                SetPixel DIS.hdc, Y, mnuHeight - X, nColor
            Next Y
        Next X
        'that rotation was simple.
        'I don't know why the msdn article was so complex.
        
        'remove the caption picture from the user form
        .Cls
        'Hopefully this operation was so fast that you did'nt see it happen.
     End With

End Sub

Public Sub MenuPopUp()
    'create the menu
    hMenu = CreatePopupMenu()
    hSubmenu = CreatePopupMenu()
    
    AppendMenu hMenu, MF_OWNERDRAW Or MF_DISABLED, 1000, 0& 'SideBar
    AppendMenu hMenu, MF_POPUP Or MF_MENUBARBREAK, hSubmenu, "Menu Item1"
    AppendMenu hMenu, MF_GRAYED, 1200, "Menu Item2"
    AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hMenu, chkMnuFlags(2), 1300, "Menu Item3"
    AppendMenu hMenu, 0&, 1400, "Menu Item4"
    AppendMenu hMenu, 0&, 1500, "Menu Item5"
    AppendMenu hMenu, 0&, 1600, "Menu Item6"
'uncomment the following 3 lines to show that the sidebar grows with our menu
'    AppendMenu hMenu, 0&, 1700, "Menu Item7"
'    AppendMenu hMenu, 0&, 1800, "Menu Item8"
'    AppendMenu hMenu, 0&, 1900, "Menu Item9"
    AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hMenu, 0&, 2000, "Exit"
    
    AppendMenu hSubmenu, chkMnuFlags(0), 1101, "SubMenu Item1"
    AppendMenu hSubmenu, chkMnuFlags(1), 1102, "SubMenu Item2"
    
 End Sub

Public Sub MenuTrack(frm As Form)
    
    GetCursorPos MP
    
    sMenu = TrackPopupMenu(hMenu, TPM_RETURNCMD, MP.X, MP.Y, 0, frm.hwnd, 0&)
    'check for clicks
    Select Case sMenu
        Case 1101
            If chkMnuFlags(0) = 0 Then
                chkMnuFlags(0) = MFT_RADIOCHECK Or MF_CHECKED
                chkMnuFlags(1) = 0&
            End If
        Case 1102
            If chkMnuFlags(1) = 0 Then
                chkMnuFlags(1) = MFT_RADIOCHECK Or MF_CHECKED
                chkMnuFlags(0) = 0&
            End If
        Case 1300
            If chkMnuFlags(2) = 0 Then
                chkMnuFlags(2) = MF_CHECKED
            Else
                chkMnuFlags(2) = 0&
            End If
        Case 2000
            UnHook
            Unload AppForm
            Exit Sub
        Case Else
            'Microsoft say's you should always have a Case Else
            'so here it is.
    End Select
    
    'update checked menu items
    ModifyMenu hMenu, 1101, chkMnuFlags(0), 1101, "Child1 Of Menu Item1"
    ModifyMenu hMenu, 1102, chkMnuFlags(1), 1102, "Child2 Of Menu Item1"
    ModifyMenu hMenu, 1300, chkMnuFlags(2), 1300, "Menu Item3"
    
    If sMenu <> 0 Then frm.Print sMenu  'just to show a response to clicking.
    
End Sub
