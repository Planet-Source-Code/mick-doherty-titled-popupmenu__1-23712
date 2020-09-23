Attribute VB_Name = "FormHook"
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, _
        ByVal hwnd As Long, _
        ByVal MSG As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

Const GWL_WNDPROC = -4

Const WM_DRAWITEM = &H2B
Const WM_MEASUREITEM = &H2C
Const WM_INITMENU = &H116
Const WM_INITMENUPOPUP = &H117

Global lpPrevWndProc As Long
Global ghw As Long
Public AppForm As Form

Public Sub Hook(frm As Form)
    Set AppForm = frm
    ghw = frm.hwnd
    lpPrevWndProc = SetWindowLong(ghw, GWL_WNDPROC, AddressOf WindowProc)
    'Set initial states of checked menuitems
    chkMnuFlags(0) = MFT_RADIOCHECK Or MF_CHECKED
    chkMnuFlags(2) = MF_CHECKED
    
    MenuPopUp       'PopMenu MenuPopUp()
End Sub

Public Sub UnHook()
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(ghw, GWL_WNDPROC, lpPrevWndProc)
    DestroyMenu hMenu
End Sub

Function WindowProc(ByVal hwnd As Long, _
            ByVal uMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long

    On Error Resume Next
    
    Select Case uMsg
        
        Case WM_MEASUREITEM 'lParam here is a pointer to a MeasureItemStruct
            MeasureMenu lParam  'PopMenu MeasureMenu()
            WindowProc = 0
        
        Case WM_DRAWITEM    'lParam here is a pointer to a DrawItemStruct
            DrawMenu lParam     'PopMenu DrawMenu()
            WindowProc = 0
        
        Case Else
            WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    
    End Select

End Function
