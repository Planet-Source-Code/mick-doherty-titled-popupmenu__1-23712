VERSION 5.00
Begin VB.Form frmPopupMenu 
   AutoRedraw      =   -1  'True
   Caption         =   "Titled Menu Demo"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   210
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.ScaleMode = vbPixels ' API works in Pixels
    Hook Me    'FormHook Hook()
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        MenuTrack Me 'PopMenu MenuTrack()
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook     'FormHook UnHook()
    DestroyMenu hMenu
End Sub

