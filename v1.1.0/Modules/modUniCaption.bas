Attribute VB_Name = "modUniCaption"
Option Explicit

Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC

Public Property Get CaptionW(ByVal hwnd As Long) As String
    Dim lngLen As Long
    lngLen = DefWindowProcW(hwnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    If lngLen Then
        CaptionW = Space$(lngLen)
        DefWindowProcW hwnd, WM_GETTEXT, lngLen + 1, StrPtr(CaptionW)
    End If
End Property

Public Property Let CaptionW(ByVal hwnd As Long, ByVal NewValue As String)
    DefWindowProcW hwnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue)
End Property
