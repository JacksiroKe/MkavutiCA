VERSION 5.00
Begin VB.UserControl xpButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   84
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   134
   ToolboxBitmap   =   "xpButton.ctx":0000
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "xpButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type RECT
    left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1 Or DT_WORDBREAK
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNTEXT = 18
Private Const RGN_DIFF = 4
Private Const PS_SOLID = 0
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()
Private He As Long
Private Wi As Long
Private BackC As Long
Private ForeC As Long
Private elTex As String
Private rc As RECT, rc2 As RECT, rc3 As RECT
Private rgnNorm As Long
Private LastButton As Byte, LastKeyDown As Byte
Private isEnabled As Boolean
Private hasFocus As Boolean, showFocusR As Boolean
Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long, cTextO As Long
Private lastStat As Byte, TE As String
Private isOver As Boolean
Private Sub OverTimer_Timer()
    Dim Pt As POINTAPI
    GetCursorPos Pt
    If UserControl.hWnd <> WindowFromPoint(Pt.X, Pt.Y) Then
        OverTimer.Enabled = False
        isOver = False
        Call Redraw(0, True)
        RaiseEvent MouseOut
    End If
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call UserControl_Click
End Sub
Private Sub UserControl_Click()
    If (LastButton = 1) And (isEnabled = True) Then
        Call Redraw(0, True)
        UserControl.Refresh
        RaiseEvent Click
    End If
End Sub
Private Sub UserControl_DblClick()
    If LastButton = 1 Then
        Call UserControl_MouseDown(1, 1, 1, 1)
    End If
End Sub
Private Sub UserControl_GotFocus()
    hasFocus = True
    Call Redraw(lastStat, True)
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    LastKeyDown = KeyCode
    If KeyCode = 32 Then
        Call UserControl_MouseDown(1, 1, 1, 1)
    ElseIf (KeyCode = 39) Or (KeyCode = 40) Then
        SendKeys "{Tab}"
    ElseIf (KeyCode = 37) Or (KeyCode = 38) Then
        SendKeys "+{Tab}"
    End If
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If (KeyCode = 32) And (LastKeyDown = 32) Then
        Call UserControl_MouseUp(1, 1, 1, 1)
        LastButton = 1
        Call UserControl_Click
    End If
End Sub
Private Sub UserControl_LostFocus()
    hasFocus = False
    Call Redraw(lastStat, True)
End Sub
Private Sub UserControl_Initialize()
    LastButton = 1
    Call SetColors
End Sub
Private Sub UserControl_InitProperties()
    isEnabled = True
    showFocusR = True
    elTex = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
    BackC = GetSysColor(COLOR_BTNFACE)
    ForeC = GetSysColor(COLOR_BTNTEXT)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    LastButton = Button
    If Button <> 2 Then Call Redraw(2, False)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If X < 0 Or Y < 0 Or X > Wi Or Y > He Then
            Call Redraw(0, False)
        Else
            If (Button = 0) And (isOver = False) Then
                OverTimer.Enabled = True
                isOver = True
                RaiseEvent MouseOver
                Call Redraw(0, True)
            ElseIf Button = 1 Then
                Call Redraw(2, False)
            End If
        End If
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button <> 2 Then Call Redraw(0, False)
End Sub
Public Property Get Caption() As String
    Caption = elTex
End Property
Public Property Let Caption(ByVal NewValue As String)
    elTex = NewValue
    Call SetAccessKeys
    Call CalculEspaceTexte
    Call Redraw(0, True)
    PropertyChanged "TX"
End Property
Public Property Get Enabled() As Boolean
    Enabled = isEnabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
    isEnabled = NewValue
    Call Redraw(0, True)
    UserControl.Enabled = isEnabled
    PropertyChanged "ENAB"
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByRef newFont As Font)
    Set UserControl.Font = newFont
    Call CalculEspaceTexte
    Call Redraw(0, True)
    PropertyChanged "FONT"
End Property
Public Property Get FontBold() As Boolean
    FontBold = UserControl.FontBold
End Property
Public Property Let FontBold(ByVal NewValue As Boolean)
    UserControl.FontBold = NewValue
    Call CalculEspaceTexte
    Call Redraw(0, True)
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property
Public Property Let FontItalic(ByVal NewValue As Boolean)
    UserControl.FontItalic = NewValue
    Call CalculEspaceTexte
    Call Redraw(0, True)
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = UserControl.FontUnderline
End Property
Public Property Let FontUnderline(ByVal NewValue As Boolean)
    UserControl.FontUnderline = NewValue
    Call CalculEspaceTexte
    Call Redraw(0, True)
End Property
Public Property Get FontSize() As Integer
    FontSize = UserControl.FontSize
End Property
Public Property Let FontSize(ByVal NewValue As Integer)
    UserControl.FontSize = NewValue
    Call CalculEspaceTexte
    Call Redraw(0, True)
End Property
Public Property Get FontName() As String
    FontName = UserControl.FontName
End Property
Public Property Let FontName(ByVal NewValue As String)
    UserControl.FontName = NewValue
    Call CalculEspaceTexte
    Call Redraw(0, True)
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal newPointer As MousePointerConstants)
    UserControl.MousePointer = newPointer
    PropertyChanged "MPTR"
End Property
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal newIcon As StdPicture)
    On Local Error Resume Next
    Set UserControl.MouseIcon = newIcon
    PropertyChanged "MICON"
End Property
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
Private Sub UserControl_Resize()
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    GetClientRect UserControl.hWnd, rc3: InflateRect rc3, -4, -4
    Call CalculEspaceTexte
    DeleteObject rgnNorm
    Call MakeRegion
    SetWindowRgn UserControl.hWnd, rgnNorm, True
    If He > 0 Then Call Redraw(0, True)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        elTex = .ReadProperty("TX", "")
        isEnabled = .ReadProperty("ENAB", True)
        Set UserControl.Font = .ReadProperty("FONT", UserControl.Font)
        showFocusR = .ReadProperty("FOCUSR", True)
        UserControl.MousePointer = .ReadProperty("MPTR", 0)
        Set UserControl.MouseIcon = .ReadProperty("MICON", Nothing)
    End With
    UserControl.Enabled = isEnabled
    Call SetColors
    Call SetAccessKeys
    Call Redraw(0, False)
End Sub
Private Sub UserControl_Terminate()
    DeleteObject rgnNorm
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("TX", elTex)
        Call .WriteProperty("ENAB", isEnabled)
        Call .WriteProperty("FONT", UserControl.Font)
        Call .WriteProperty("FOCUSR", showFocusR)
        Call .WriteProperty("MPTR", UserControl.MousePointer)
        Call .WriteProperty("MICON", UserControl.MouseIcon)
    End With
End Sub
Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
    If Force = False Then
        If (curStat = lastStat) And (TE = elTex) Then Exit Sub
    End If
    If He = 0 Then Exit Sub
    lastStat = curStat
    TE = elTex
    Dim i As Long, stepXP1 As Single, XPface As Long
    With UserControl
        .Cls
        DrawRectangle 0, 0, Wi, He, cFace
        If isEnabled = True Then
            'set font color
            If isOver Then
                SetTextColor .hDC, cTextO
            Else
                SetTextColor .hDC, cText
            End If
            If curStat = 0 Then
                stepXP1 = 25 / He
                XPface = ShiftColor(cFace, &H30, True)
                For i = 1 To He
                    DrawLine 0, i, Wi, i, ShiftColor(XPface, -stepXP1 * i, True)
                Next
                DrawText .hDC, elTex, Len(elTex), rc, DT_CENTER
                DrawRectangle 0, 0, Wi, He, &H733C00, True
                mSetPixel 1, 1, &H7B4D10
                mSetPixel 1, He - 2, &H7B4D10
                mSetPixel Wi - 2, 1, &H7B4D10
                mSetPixel Wi - 2, He - 2, &H7B4D10
                If isOver Then
                    DrawRectangle 1, 2, Wi - 2, He - 4, &H31B2FF, True
                    DrawLine 2, He - 2, Wi - 2, He - 2, &H96E7&
                    DrawLine 2, 1, Wi - 2, 1, &HCEF3FF
                    DrawLine 1, 2, Wi - 1, 2, &H8CDBFF
                    DrawLine 2, 3, 2, He - 3, &H6BCBFF
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, &H6BCBFF
                ElseIf ((hasFocus Or Ambient.DisplayAsDefault) And showFocusR) Then
                    DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                    DrawLine 2, He - 2, Wi - 2, He - 2, &HEF826B
                    DrawLine 2, 1, Wi - 2, 1, &HFFE7CE
                    DrawLine 1, 2, Wi - 1, 2, &HF7D7BD
                    DrawLine 2, 3, 2, He - 3, &HF0D1B5
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5
                Else
                    DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPface, -&H30, True)
                    DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPface, -&H20, True)
                    DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPface, -&H24, True)
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, ShiftColor(XPface, -&H18, True)
                    DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPface, &H10, True)
                    DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPface, &HA, True)
                    DrawLine 1, 2, 1, He - 2, ShiftColor(XPface, -&H5, True)
                    DrawLine 2, 3, 2, He - 3, ShiftColor(XPface, -&HA, True)
                End If
            ElseIf curStat = 2 Then
                stepXP1 = 25 / He
                XPface = ShiftColor(cFace, &H30, True)
                XPface = ShiftColor(XPface, -32, True)
                For i = 1 To He
                    DrawLine 0, He - i, Wi, He - i, ShiftColor(XPface, -stepXP1 * i, True)
                Next
                SetTextColor .hDC, cText
                DrawText .hDC, elTex, Len(elTex), rc2, DT_CENTER
                DrawRectangle 0, 0, Wi, He, &H733C00, True
                mSetPixel 1, 1, &H7B4D10
                mSetPixel 1, He - 2, &H7B4D10
                mSetPixel Wi - 2, 1, &H7B4D10
                mSetPixel Wi - 2, He - 2, &H7B4D10
                DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPface, &H10, True)
                DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPface, &HA, True)
                DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPface, &H5, True)
                DrawLine Wi - 3, 3, Wi - 3, He - 3, XPface
                DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPface, -&H20, True)
                DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPface, -&H18, True)
                DrawLine 1, 2, 1, He - 2, ShiftColor(XPface, -&H20, True)
                DrawLine 2, 2, 2, He - 2, ShiftColor(XPface, -&H16, True)
            End If
        Else
            XPface = ShiftColor(cFace, &H30, True)
            DrawRectangle 0, 0, Wi, He, ShiftColor(XPface, -&H18, True)
            SetTextColor .hDC, ShiftColor(XPface, -&H68, True)
            DrawText .hDC, elTex, Len(elTex), rc, DT_CENTER
            DrawRectangle 0, 0, Wi, He, ShiftColor(XPface, -&H54, True), True
            mSetPixel 1, 1, ShiftColor(XPface, -&H48, True)
            mSetPixel 1, He - 2, ShiftColor(XPface, -&H48, True)
            mSetPixel Wi - 2, 1, ShiftColor(XPface, -&H48, True)
            mSetPixel Wi - 2, He - 2, ShiftColor(XPface, -&H48, True)
        End If
    End With
End Sub
Private Sub DrawRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
    Dim bRect As RECT
    Dim hBrush As Long
    Dim Ret As Long
    bRect.left = X
    bRect.Top = Y
    bRect.Right = X + Width
    bRect.Bottom = Y + Height
    hBrush = CreateSolidBrush(Color)
    If OnlyBorder = False Then
        Ret = FillRect(UserControl.hDC, bRect, hBrush)
    Else
        Ret = FrameRect(UserControl.hDC, bRect, hBrush)
    End If
    Ret = DeleteObject(hBrush)
End Sub
Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    Dim Pt As POINTAPI
    Dim oldPen As Long, hPen As Long
    With UserControl
        hPen = CreatePen(PS_SOLID, 1, Color)
        oldPen = SelectObject(.hDC, hPen)
        MoveToEx .hDC, X1, Y1, Pt
        LineTo .hDC, X2, Y2
        SelectObject .hDC, oldPen
        DeleteObject hPen
    End With
End Sub
Private Sub mSetPixel(ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    Call SetPixel(UserControl.hDC, X, Y, Color)
End Sub
Private Sub SetColors()
    cFace = &HC0C0C0
    cShadow = &H808080
    cLight = &HDFDFDF
    cDarkShadow = &H0
    cHighLight = &HFFFFFF
    cText = &H0
    cTextO = cText
End Sub
Private Sub MakeRegion()
    Dim rgn1 As Long, rgn2 As Long
    DeleteObject rgnNorm
    rgnNorm = CreateRectRgn(0, 0, Wi, He)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    rgn1 = CreateRectRgn(0, 0, 2, 1)
    CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, He, 2, He - 1)
    CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
    CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
    CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, 1, 1, 2)
    CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
    CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
    CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
    DeleteObject rgn1
    rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
    CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    DeleteObject rgn1
    DeleteObject rgn2
End Sub
Private Sub SetAccessKeys()
    Dim ampersandPos As Long
    If Len(elTex) > 1 Then
        ampersandPos = InStr(1, elTex, "&", vbTextCompare)
        If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
            If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then
                UserControl.AccessKeys = LCase$(Mid$(elTex, ampersandPos + 1, 1))
            Else
                ampersandPos = InStr(ampersandPos + 2, elTex, "&", vbTextCompare)
                If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then
                    UserControl.AccessKeys = LCase$(Mid$(elTex, ampersandPos + 1, 1))
                Else
                    UserControl.AccessKeys = ""
                End If
            End If
        Else
            UserControl.AccessKeys = ""
        End If
    Else
        UserControl.AccessKeys = ""
    End If
End Sub
Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
    Dim Red As Long, Blue As Long, Green As Long
    If isXP = False Then
        Blue = ((Color \ &H10000) Mod &H100) + Value
    Else
        Blue = ((Color \ &H10000) Mod &H100)
        Blue = Blue + ((Blue * Value) \ &HC0)
    End If
    Green = ((Color \ &H100) Mod &H100) + Value
    Red = (Color And &HFF) + Value
    If Red < 0 Then
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    If Green < 0 Then
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If
    ShiftColor = RGB(Red, Green, Blue)
End Function
Private Sub CalculEspaceTexte()
    rc2.left = 1: rc2.Right = Wi - 2: rc2.Top = 0: rc2.Bottom = He - 2
    DrawText UserControl.hDC, elTex, Len(elTex), rc2, DT_CALCRECT Or DT_WORDBREAK
    CopyRect rc, rc2: OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom) \ 2
    CopyRect rc2, rc: OffsetRect rc2, 1, 1
End Sub
Public Sub Refresh()
    Call Redraw(lastStat, True)
End Sub
