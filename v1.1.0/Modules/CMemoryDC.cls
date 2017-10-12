VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemoryDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*******************************************************************************
'*    Component   : CMemoryDC
'*    Description : Memory device contex drawing class used for flicker-free
'*                  drawing.
'*    Credits     : This class contains some functions which I(Andrea Batina)
'                   found on internet.
'*                  If you see your function in here please notify me(Andrea Batina)
'*                  so that I can put your name on it.
'*
'*    Copyright   : Copyright © 2004 Revelatek. All rights reserved.
'*******************************************************************************

Option Explicit

' Base error number constant
Private Const ERRBASE = 5000

' --Private/Public Type Definitions
Private Type POINTAPI
    x                   As Long
    y                   As Long
End Type

' --For gradient subs
Public Enum GradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum

' --A trick to preserve casing of enums when typing in IDE
#If False Then
Private gdHorizontal, gdVertical, gdDownwardDiagonal, gdUpwardDiagonal
#End If

Private Type DRAWTEXTPARAMS
    cbSize              As Long
    iTabLength          As Long
    iLeftMargin         As Long
    iRightMargin        As Long
    uiLengthDrawn       As Long
End Type

Private Type TEXTMETRIC
    tmHeight            As Long
    tmAscent            As Long
    tmDescent           As Long
    tmInternalLeading   As Long
    tmExternalLeading   As Long
    tmAveCharWidth      As Long
    tmMaxCharWidth      As Long
    tmWeight            As Long
    tmOverhang          As Long
    tmDigitizedAspectX  As Long
    tmDigitizedAspectY  As Long
    tmFirstChar         As Byte
    tmLastChar          As Byte
    tmDefaultChar       As Byte
    tmBreakChar         As Byte
    tmItalic            As Byte
    tmUnderlined        As Byte
    tmStruckOut         As Byte
    tmPitchAndFamily    As Byte
    tmCharSet           As Byte
End Type

Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(1 To 32) As Byte
End Type

Private Type SIZEAPI
    cX                  As Long
    cY                  As Long
End Type
Private Type PointSng   ' Internal Point structure
    x                   As Single   ' Uses Singles for more precision.
    y                   As Single
End Type

Private Type RGBType
    r                   As Byte
    g                   As Byte
    B                   As Byte
    A                   As Byte
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

'  for gradient painting and bitmap tiling
Private Type BITMAPINFOHEADER
   biSize               As Long
   biWidth              As Long
   biHeight             As Long
   biPlanes             As Integer
   biBitCount           As Integer
   biCompression        As Long
   biSizeImage          As Long
   biXPelsPerMeter      As Long
   biYPelsPerMeter      As Long
   biClrUsed            As Long
   biClrImportant       As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0

Private Type PictDesc
    Size                As Long
    Type                As Long
    hBmpOrIcon          As Long
    hPal                As Long
End Type

' --Private/Public Enum Definitions
Public Enum EMemDCDrawText
    DT_LEFT = &H0               ' Aligns text to the left.
    DT_TOP = &H0                ' Justifies the text to the top of the rectangle.
    DT_CENTER = &H1             ' Centers text horizontally in the rectangle.
    DT_RIGHT = &H2              ' Aligns text to the right.
    DT_VCENTER = &H4            ' Centers text vertically. This value is used only with the DT_SINGLELINE value.
    DT_BOTTOM = &H8             ' Justifies the text to the bottom of the rectangle. This value is used only with the DT_SINGLELINE value.
    DT_WORDBREAK = &H10         ' Breaks words. Lines are automatically broken between words if a word would extend past the edge of the rectangle specified by the lpRect parameter. A carriage return-line feed sequence also breaks the line.<br>If this is not specified, output is on one line.
    DT_SINGLELINE = &H20        ' Displays text on a single line only. Carriage returns and line feeds do not break the line.
    DT_EXPANDTABS = &H40        ' Expands tab characters. The default number of characters per tab is eight. The DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS, and DT_END_ELLIPSIS values cannot be used with the DT_EXPANDTABS value.
    DT_TABSTOP = &H80           ' Sets tab stops. Bits 15–8 (high-order byte of the low-order word) of the uFormat parameter specify the number of characters for each tab. The default number of characters per tab is eight. The DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, and DT_NOPREFIX values cannot be used with the DT_TABSTOP value.
    DT_NOCLIP = &H100           ' Draws without clipping. DrawText is somewhat faster when DT_NOCLIP is used.
    DT_EXTERNALLEADING = &H200  ' Includes the font external leading in line height. Normally, external leading is not included in the height of a line of text.
    DT_CALCRECT = &H400         ' Determines the width and height of the rectangle. If there are multiple lines of text, DrawText uses the width of the rectangle pointed to by the lpRect parameter and extends the base of the rectangle to bound the last line of text. If the largest word is wider than the rectangle, the width is expanded. If the text is less than the width of the rectangle, the width is reduced. If there is only one line of text, DrawText modifies the right side of the rectangle so that it bounds the last character in the line. In either case, DrawText returns the height of the formatted text but does not draw the text.
    DT_NOPREFIX = &H800         ' Turns off processing of prefix characters. Normally, DrawText interprets the mnemonic-prefix character & as a directive to underscore the character that follows, and the mnemonic-prefix characters && as a directive to print a single &. By specifying DT_NOPREFIX, this processing is turned off
    DT_INTERNAL = &H1000        ' Uses the system font to calculate text metrics.
    DT_EDITCONTROL = &H2000     ' Duplicates the text-displaying characteristics of a multiline edit control. Specifically, the average character width is calculated in the same manner as for an edit control, and the function does not display a partially visible last line.
    DT_PATH_ELLIPSIS = &H4000   ' For displayed text, replaces characters in the middle of the string with ellipses so that the result fits in the specified rectangle. If the string contains backslash (\) characters, DT_PATH_ELLIPSIS preserves as much as possible of the text after the last backslash.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_END_ELLIPSIS and DT_WORD_ELLIPSIS.
    DT_END_ELLIPSIS = &H8000    ' For displayed text, if the end of a string does not fit in the rectangle, it is truncated and ellipses are added. If a word that is not at the end of the string goes beyond the limits of the rectangle, it is truncated without ellipses.<br>The string is not modified unless the DT_MODIFYSTRING flag is specified.<br>Compare with DT_PATH_ELLIPSIS and DT_WORD_ELLIPSIS.
    DT_MODIFYSTRING = &H10000   ' Modifies the specified string to match the displayed text. This value has no effect unless DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is specified.
    DT_RTLREADING = &H20000     ' Layout in right-to-left reading order for bi-directional text when the font selected into the hdc is a Hebrew or Arabic font. The default reading order for all text is left-to-right.
    DT_WORD_ELLIPSIS = &H40000  ' Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
End Enum

#If False Then
Private DT_LEFT, DT_TOP, DT_CENTER, DT_RIGHT, DT_VCENTER, DT_BOTTOM, DT_WORDBREAK, DT_SINGLELINE, DT_EXPANDTABS, DT_TABSTOP, DT_NOCLIP, _
        DT_EXTERNALLEADING, DT_CALCRECT, DT_NOPREFIX, DT_INTERNAL, DT_EDITCONTROL, DT_PATH_ELLIPSIS, DT_END_ELLIPSIS, DT_MODIFYSTRING, _
        DT_RTLREADING, DT_WORD_ELLIPSIS  ':) Line inserted by Formatter
#End If


' --Private/Public Win32 API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal nBkMode As Long) As Long
Private Declare Function APIFillRect Lib "user32" Alias "FillRect" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function APIBitBlt Lib "gdi32" Alias "BitBlt" (ByVal hdcDest As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function ApiDrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hDc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextExtentPointA Lib "gdi32" (ByVal hDc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Private Declare Function GetTextExtentPointW Lib "gdi32" (ByVal hDc As Long, ByVal lpszString As Long, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Private Declare Function APIGetPixel Lib "gdi32" Alias "GetPixel" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function APISetPixel Lib "gdi32" Alias "SetPixel" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long


' --Private/Public Constant Declarations
' Pen Style Messages
Private Const PS_SOLID = 0
Private Const PS_DOT = 2

' Draw State Messages
Private Const DST_BITMAP = &H4
Private Const DST_COMPLEX = &H0
Private Const DST_ICON = &H3
Private Const DSS_NORMAL = &H0
Private Const DSS_MONO = &H80

' Misc Messages
Private Const LOGPIXELSY = 90
Private Const FW_NORMAL = 400
Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180


' --Private/Public Variable Declarations
Private m_lhDC              As Long     ' Memory dc handle
Private m_lWidth            As Long     ' Memory dc width
Private m_lHeight           As Long     ' Memory dc height
Private m_hBmp              As Long     ' Memory dc new compatible bitmap
Private m_hBmpOld           As Long     ' Memory dc old compatible bitmap
Private m_lOrginalFont      As Long     ' Memory dc original font
Private m_lMemoryFont       As Long     ' Memory dc new font

'********************************************************************
'* Name: AlphaBlend
'* Description: Returns alpha blend of two specified colors.
'********************************************************************

Public Function AlphaBlend(ByVal lColorOne As Long, ByVal lColorTwo As Long, ByVal lAlpha As Long) As OLE_COLOR

Dim tClrFore         As RGBType
Dim tClrBack         As RGBType

' Translate both colors

    lColorOne = TranslateColor(lColorOne)
    lColorTwo = TranslateColor(lColorTwo)
    ' Split second color into RGB elements
    With tClrBack
        .r = GetRGB(lColorTwo, 1)
        .g = GetRGB(lColorTwo, 2)
        .B = GetRGB(lColorTwo, 3)
    End With
    With tClrFore
        ' Split the first color also
        .r = GetRGB(lColorOne, 1)
        .g = GetRGB(lColorOne, 2)
        .B = GetRGB(lColorOne, 3)
        ' Alpha blend each color item
        .r = (.r * lAlpha + tClrBack.r * (255 - lAlpha)) / 255
        .g = (.g * lAlpha + tClrBack.g * (255 - lAlpha)) / 255
        .B = (.B * lAlpha + tClrBack.B * (255 - lAlpha)) / 255
    End With
    ' Return blended color
    AlphaBlend = RGB(tClrFore.r, tClrFore.g, tClrFore.B)

End Function

'********************************************************************
'* Name: BitBlt
'* Description: Transfers image from memory dc to hdcDest dc.
'********************************************************************

Public Sub BitBlt(ByVal hdcDest As Long, _
                  Optional ByVal xDest As Long, _
                  Optional ByVal yDest As Long, _
                  Optional ByVal lWidth As Long, _
                  Optional ByVal lHeight As Long, _
                  Optional ByVal xSrc As Long, _
                  Optional ByVal ySrc As Long, _
                  Optional ByVal dwRop As RasterOpConstants = vbSrcCopy, _
                  Optional ByVal lhdcSrc As Long = 0)

    On Error GoTo PROC_ERR_BitBlt

    If Not IsCreated Then
        Exit Sub
    End If

    ' Setup properties
    If lWidth = 0 Then
        lWidth = m_lWidth
    End If
    
    If lHeight = 0 Then
        lHeight = m_lHeight
    End If
    ' BitBlt
    If lhdcSrc <> 0 Then
        APIBitBlt hdcDest, xDest, yDest, lWidth, lHeight, lhdcSrc, xSrc, ySrc, dwRop
    Else
        APIBitBlt hdcDest, xDest, yDest, lWidth, lHeight, m_lhDC, xSrc, ySrc, dwRop
    End If

PROC_EXIT:

Exit Sub

PROC_ERR_BitBlt:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.BitBlt", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

'********************************************************************
'* Name: Class_Terminate
'* Description: Class object termination.
'********************************************************************

Private Sub Class_Terminate()

    On Error Resume Next
    pDestroyMemDC

End Sub

'********************************************************************
'* Name: Cls
'* Description: Clear memory dc.
'********************************************************************

Public Sub Cls(Optional ByVal lColor As Long = -1)

    FillRect 0, 0, Width, Height, lColor

End Sub

'********************************************************************
'* Name: Draw3DRect
'* Description: Draw 3D rectangle.
'********************************************************************

Public Sub Draw3DRect( _
                      ByVal lLeft As Long, _
                      ByVal lTop As Long, _
                      ByVal lRight As Long, _
                      ByVal lBottom As Long, _
                      ByVal oTopLeftColor As OLE_COLOR, _
                      ByVal oBottomRightColor As OLE_COLOR)

    On Error GoTo PROC_ERR_Draw3DRect

    If Not IsCreated() Then
        Exit Sub
    End If

    Dim hPen As Long
    Dim hPenOld As Long
    Dim tP As POINTAPI
    Dim rcItem As RECT

    ' Setup rectangle dimensions
    With rcItem
        .Left = lLeft
        .Top = lTop
        .Right = lRight
        .Bottom = lBottom
    End With

    ' Create pen with specified color
    hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
    ' Save old pen
    hPenOld = SelectObject(m_lhDC, hPen)
    ' Draw top-left of rectangle
    MoveToEx m_lhDC, rcItem.Left, rcItem.Bottom - 1, tP
    LineTo m_lhDC, rcItem.Left, rcItem.Top
    LineTo m_lhDC, rcItem.Right - 1, rcItem.Top
    ' Restore old pen
    SelectObject m_lhDC, hPenOld
    ' Delete pen
    DeleteObject hPen
    ' If there is right border
    If rcItem.Left <> rcItem.Right Then
        ' Create pen
        hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
        ' Save old pen
        hPenOld = SelectObject(m_lhDC, hPen)
        ' Draw bottom-right of rectangle
        LineTo m_lhDC, rcItem.Right - 1, rcItem.Bottom - 1
        LineTo m_lhDC, rcItem.Left - 1, rcItem.Bottom - 1
        ' Restore old pen
        SelectObject m_lhDC, hPenOld
        ' Delete pen
        DeleteObject hPen
    End If

PROC_EXIT:

Exit Sub

PROC_ERR_Draw3DRect:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.Draw3DRect", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

Public Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long, Optional lHdc As Long = 0)

'****************************************************************************
'* draw lines
'****************************************************************************

Dim pt               As POINTAPI
Dim hPen             As Long
Dim hPenOld          As Long

    If Not IsCreated Then
        Exit Sub
    End If
    
    If lHdc <> 0 Then
        hPen = CreatePen(0, 1, TranslateColor(Color))
        hPenOld = SelectObject(lHdc, hPen)
        MoveToEx lHdc, X1, Y1, pt
        LineTo lHdc, X2, Y2
        SelectObject lHdc, hPenOld
        DeleteObject hPen
    End If
    
    hPen = CreatePen(0, 1, TranslateColor(Color))
    hPenOld = SelectObject(m_lhDC, hPen)
    MoveToEx m_lhDC, X1, Y1, pt
    LineTo m_lhDC, X2, Y2
    SelectObject m_lhDC, hPenOld
    DeleteObject hPen

End Sub

Public Sub DrawGradient(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal GradientDirection As GradientDirectionCts)

'****************************************************************************
'* Draws very fast Gradient in four direction.                              *
'* Author: Carles P.V (Gradient Master)                                     *
'****************************************************************************

Dim uBIH             As BITMAPINFOHEADER
Dim lBits()          As Long
Dim lGrad()          As Long

Dim r1               As Long
Dim g1               As Long
Dim b1               As Long
Dim r2               As Long
Dim g2               As Long
Dim b2               As Long
Dim dR               As Long
Dim dG               As Long
Dim dB               As Long

Dim Scan             As Long
Dim i                As Long
Dim iEnd             As Long
Dim iOffset          As Long
Dim J                As Long
Dim jEnd             As Long
Dim iGrad            As Long

Dim lTmpColor        As Long

    '-- A minor check

   If (Width < 1 Or Height < 1) Then
        Exit Sub
    End If
    
    ' --I have swapped colors
    ' --as Andrea was using different DrawGradient routine
    ' --and was interchanging the angle :-)
    '****************************
    lTmpColor = Color2          '
    Color2 = Color1             '
    Color1 = lTmpColor          '
    '****************************
    
    '-- Decompose colors
    Color1 = Color1 And &HFFFFFF
    r1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    g1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    b1 = Color1 Mod &H100&
    Color2 = Color2 And &HFFFFFF
    r2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    g2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    b2 = Color2 Mod &H100&

    '-- Get color distances
    dR = r2 - r1
    dG = g2 - g1
    dB = b2 - b1

    '-- Size gradient-colors array
    Select Case GradientDirection
    Case [gdHorizontal]
        ReDim lGrad(0 To Width - 1)
    Case [gdVertical]
        ReDim lGrad(0 To Height - 1)
    Case Else
        ReDim lGrad(0 To Width + Height - 2)
    End Select

    '-- Calculate gradient-colors
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (g1 \ 2 + g2 \ 2) + 65536 * (r1 \ 2 + r2 \ 2)
    Else
        For i = 0 To iEnd
            lGrad(i) = b1 + (dB * i) \ iEnd + 256 * (g1 + (dG * i) \ iEnd) + 65536 * (r1 + (dR * i) \ iEnd)
        Next i
    End If

    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width

    '-- Render gradient DIB
    Select Case GradientDirection

    Case [gdHorizontal]

        For J = 0 To jEnd
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(i - iOffset)
            Next i
            iOffset = iOffset + Scan
        Next J

    Case [gdVertical]

        For J = jEnd To 0 Step -1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(J)
            Next i
            iOffset = iOffset + Scan
        Next J

    Case [gdDownwardDiagonal]

        iOffset = jEnd * Scan
        For J = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(iGrad)
                iGrad = iGrad + 1
            Next i
            iOffset = iOffset - Scan
            iGrad = J
        Next J

    Case [gdUpwardDiagonal]

        iOffset = 0
        For J = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(iGrad)
                iGrad = iGrad + 1
            Next i
            iOffset = iOffset + Scan
            iGrad = J
        Next J
    End Select

    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With

    '-- Paint it!
    StretchDIBits m_lhDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy

End Sub

'********************************************************************
'* Name: DrawPicture
'* Description: Draw specified picture on memory dc.
'********************************************************************

Public Sub DrawPicture(ByRef m_PictureF As StdPicture, _
                       ByVal x As Long, _
                       ByVal y As Long, _
                       Optional lW As Long = 16, _
                       Optional lH As Long = 16, _
                       Optional bShadow As Boolean = False, _
                       Optional lShadowColor As Long = -1, _
                       Optional lMaskColor As Long = -1)

Dim lFlags As Long
Dim hBrush As Long

    'On Error GoTo PROC_ERR_DrawPicture

    If Not IsCreated() Then
        Exit Sub
    End If

    ' Get image type
    Select Case m_PictureF.Type
    Case vbPicTypeBitmap
        lFlags = DST_BITMAP

        Dim hdcPaint As Long, hbmOrig As Long
        hdcPaint = CreateCompatibleDC(m_lhDC)
        hbmOrig = SelectObject(hdcPaint, m_PictureF.handle)
        If lMaskColor = -1 Then
            APIBitBlt m_lhDC, x, y, lW, lH, hdcPaint, 0, 0, vbSrcCopy
        Else
            pvTransBlt m_lhDC, x, y, lW, lH, hdcPaint, 0, 0, lMaskColor
        End If
        SelectObject hdcPaint, hbmOrig
        DeleteDC hdcPaint
        Exit Sub

    Case vbPicTypeIcon
        lFlags = DST_ICON
    Case Else
        lFlags = DST_COMPLEX
    End Select
    ' If shadow is set then create shadow brush
    If bShadow Then
        If lShadowColor <> -1 Then
            hBrush = CreateSolidBrush(lShadowColor)
        Else
            hBrush = CreateSolidBrush(TranslateColor(vbButtonShadow))
        End If
    End If
    ' Draw icon
    DrawState m_lhDC, IIf(bShadow, hBrush, 0), 0, m_PictureF.handle, 0, x, y, lW, lH, lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)

    ' Delete shadow brush object
    If bShadow Then
        DeleteObject hBrush
    End If

PROC_EXIT:

Exit Sub

PROC_ERR_DrawPicture:
'    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.DrawPicture", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

'********************************************************************
'* Name: DrawText
'* Description: Draw text on memory device contex.
'********************************************************************

Public Sub DrawText( _
                    ByVal sText As String, _
                    ByRef lLeft As Long, _
                    ByRef lTop As Long, _
                    ByRef lRight As Long, _
                    ByRef lBottom As Long, _
                    Optional ByVal lFlags As EMemDCDrawText)

    On Error GoTo PROC_ERR_DrawText

    If Not IsCreated() Then
        Exit Sub
    End If

Dim rc As RECT
Dim wTextParams As DRAWTEXTPARAMS

    ' Setup text box dimensions
    With rc
        .Left = lLeft
        .Top = lTop
        .Right = lRight
        .Bottom = lBottom
    End With
    wTextParams.cbSize = Len(wTextParams)

    ' If windows 9x
    If Not pIsWinNT Then
        ApiDrawTextEx m_lhDC, sText, Len(sText), rc, lFlags, wTextParams
    Else
        DrawTextW m_lhDC, StrPtr(sText), -1, rc, lFlags Or DT_RTLREADING
    End If

    ' Return text dimensions
    With rc
        lLeft = .Left
        lTop = .Top
        lRight = .Right
        lBottom = .Bottom
    End With

PROC_EXIT:

Exit Sub

PROC_ERR_DrawText:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.DrawText", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

'********************************************************************
'* Name: FillRect
'* Description: Fill rect with specified color.
'********************************************************************

Public Sub FillRect( _
                    Optional ByVal lLeft As Long, _
                    Optional ByVal lTop As Long, _
                    Optional ByVal lRight As Long = -1, _
                    Optional ByVal lBottom As Long = -1, _
                    Optional ByVal lColor As Long = -1, _
                    Optional ByVal lHdc As Long)

Dim rc As RECT
Dim hbrFill As Long
Dim hOldBrush As Long

    On Error GoTo PROC_ERR_FillRect

    If Not IsCreated Then
        Exit Sub
    End If

    ' Create solid brush with specified color
    If lColor <> -1 Then
        hbrFill = CreateSolidBrush(TranslateColor(lColor))
        If lHdc <> 0 Then
            hOldBrush = SelectObject(lHdc, hbrFill)
        Else
            hOldBrush = SelectObject(m_lhDC, hbrFill)
        End If
    End If
    
    ' Setup rectangle dimensions
    With rc
        .Left = lLeft
        .Top = lTop
        .Right = IIf(lRight < lLeft, Width, lRight)
        .Bottom = IIf(lBottom < lTop, Height, lBottom)
    End With
    
    ' Fill rect
    If lHdc <> 0 Then
        APIFillRect lHdc, rc, hbrFill
    Else
        APIFillRect m_lhDC, rc, hbrFill
    End If
    
    ' Delete brush object
    If lColor <> -1 Then
        If lHdc <> 0 Then
            SelectObject lHdc, hOldBrush
        Else
            SelectObject m_lhDC, hOldBrush
        End If
        DeleteObject hbrFill
    End If

PROC_EXIT:

Exit Sub

PROC_ERR_FillRect:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.FillRect", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

Public Property Set Font(ByVal oValue As StdFont)

    On Error GoTo PROC_ERR_Font

    If Not IsCreated() Then
        Exit Property
    End If

    Dim tFont           As LOGFONT

    ' Fill LOGFONT type with new font properties
    With tFont
        CopyMemory .lfFaceName(1), ByVal oValue.Name, Len(oValue.Name) + 1
        .lfCharSet = oValue.Charset
        .lfItalic = (-oValue.Italic)
        .lfStrikeOut = (-oValue.Strikethrough)
        .lfUnderline = (-oValue.Underline)
        .lfWeight = oValue.Weight
        .lfHeight = -(oValue.Size * GetDeviceCaps(m_lhDC, LOGPIXELSY) \ 72)
    End With
    
    ' Select new font into memory dc and delete the previous one
    If m_lMemoryFont <> 0 Then
        SelectObject m_lhDC, m_lOrginalFont
        DeleteObject m_lMemoryFont
    End If
    
    m_lMemoryFont = CreateFontIndirect(tFont)
    m_lOrginalFont = SelectObject(m_lhDC, m_lMemoryFont)

PROC_EXIT:

Exit Property

PROC_ERR_Font:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.Font", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Property

'********************************************************************
'* Name: Font
'* Description: Return\sets memory dc font.
'********************************************************************

Public Property Get Font() As StdFont

    On Error GoTo PROC_ERR_Font

    If Not IsCreated() Then
        Exit Property
    End If

Dim tTM             As TEXTMETRIC
Dim sFaceName       As String * 80

    ' Get text metrics
    GetTextMetrics m_lhDC, tTM
    ' Return font name
    GetTextFace m_lhDC, 79, sFaceName
    ' Create a new font object
    Set Font = New StdFont
    ' And fill it with memory dc font properties
    With Font
        .Name = sFaceName
        .Bold = (tTM.tmWeight >= FW_NORMAL)
        .Charset = tTM.tmCharSet
        .Italic = (tTM.tmItalic <> 0)
        .Strikethrough = (tTM.tmStruckOut <> 0)
        .Underline = (tTM.tmUnderlined <> 0)
        .Weight = tTM.tmWeight
        .Size = (tTM.tmHeight - tTM.tmInternalLeading) * 72 / tTM.tmDigitizedAspectY
    End With

PROC_EXIT:

Exit Property

PROC_ERR_Font:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.Font", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Property

'********************************************************************
'* Name: ForeColor
'* Description: Return\sets text fore color.
'********************************************************************

Public Property Get ForeColor() As Long

    On Error GoTo PROC_ERR_ForeColor

    If IsCreated() Then
        ForeColor = GetTextColor(m_lhDC)
    End If

PROC_EXIT:

Exit Property

PROC_ERR_ForeColor:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.ForeColor", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Property

Public Property Let ForeColor(ByVal lValue As Long)

    On Error GoTo PROC_ERR_ForeColor

    If IsCreated() Then
        SetTextColor m_lhDC, TranslateColor(lValue)
    End If

PROC_EXIT:

Exit Property

PROC_ERR_ForeColor:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.ForeColor", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Property

'********************************************************************
'* Name: GetPixel
'* Description: Return current pixel color.
'********************************************************************

Public Function GetPixel(ByVal x As Long, ByVal y As Long, Optional lHdc As Long = -1) As Long

    On Error GoTo PROC_ERR_GetPixel

    If Not IsCreated() Then
        Exit Function
    End If

    ' If user specified from which dc to return pixel color then
    ' return it from that one otherwise from memory dc
    If lHdc <> -1 Then
        GetPixel = APIGetPixel(lHdc, x, y)
    Else
        GetPixel = APIGetPixel(m_lhDC, x, y)
    End If

PROC_EXIT:

Exit Function

PROC_ERR_GetPixel:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.GetPixel", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Function

'********************************************************************
'* Name: GetRGB
'* Description: Extract Red, Green or Blue color from RGB color.
'********************************************************************

Public Function GetRGB(lRGB As Long, lNum As Integer) As Integer

    On Error GoTo PROC_ERR_GetRGB

    If lNum > 0 And lNum < 4 And lRGB > -1 And lRGB < 16777216 Then
        GetRGB = lRGB \ 256 ^ (lNum - 1) And 255
    End If

PROC_EXIT:

Exit Function

PROC_ERR_GetRGB:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.GetRGB", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Function

'********************************************************************
'* Name: hDC
'* Description: Returns memory dc handle.
'********************************************************************

Public Property Get hDc() As Long

    hDc = m_lhDC

End Property

Public Property Let Height(ByVal lH As Long)

    If lH > m_lHeight Then
        pCreateMemDC m_lWidth, lH
    End If

End Property

'********************************************************************
'* Name: Height
'* Description: Get\Set memory dc height.
'********************************************************************

Public Property Get Height() As Long

    Height = m_lHeight

End Property

'********************************************************************
'* Name: IconToPicture
'* Description: Creates a StdPicture object from HICON handle.
'********************************************************************

Public Function IconToPicture(ByVal hIcon As Long) As IPicture

Dim oNewPic         As Picture
Dim lpPictDesc      As PictDesc
Dim aGuid(0 To 3)   As Long

    On Error Resume Next
        '--- check argument
        'If hIcon = 0 Then Exit Function':( --> replaced by:
        If hIcon = 0 Then
            Exit Function
        End If
        '--- fill struct
        With lpPictDesc
            .Size = Len(lpPictDesc)
            .Type = vbPicTypeIcon
            .hBmpOrIcon = hIcon
        End With
        '--- fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        aGuid(0) = &H7BF80980
        aGuid(1) = &H101ABF32
        aGuid(2) = &HAA00BB8B
        aGuid(3) = &HAB0C3000
        '--- do convert
        OleCreatePictureIndirect lpPictDesc, aGuid(0), True, oNewPic
        '--- success
        Set IconToPicture = oNewPic

End Function

'********************************************************************
'* Name: IsCreated
'* Description: Returns true if we already created the memory dc.
'********************************************************************

Public Function IsCreated() As Boolean

    IsCreated = m_lhDC <> 0

End Function

'********************************************************************
'* Name: LightenColor
'* Description: Lightens provided color by specified amount.
'********************************************************************

Public Function LightenColor(ByVal lClrLight As Long, ByVal lClrDark As Long, ByVal lAmount As Long) As Long

Dim lVal As Long
Dim lRed As Long, lGreen As Long, lBlue As Long

    If lAmount < 100 Then
        lVal = 100
    Else
        lVal = 1000
    End If

    lRed = (GetRGB(lClrDark, 1) * (lVal - lAmount) + lVal / 2 + GetRGB(lClrLight, 1) * lAmount) / lVal
    lGreen = (GetRGB(lClrDark, 2) * (lVal - lAmount) + lVal / 2 + GetRGB(lClrLight, 2) * lAmount) / lVal
    lBlue = (GetRGB(lClrDark, 3) * (lVal - lAmount) + lVal / 2 + GetRGB(lClrLight, 3) * lAmount) / lVal

    LightenColor = RGB(lRed, lGreen, lBlue)

End Function

'********************************************************************
'* Name: pBlendColors
'* Description: Creates an array of colors blending from Color1 to
'* Color2 in lSteps number of steps. Returns the count and fills
'* the laRetColors() array.
'* Credits: Kath-Rock
'********************************************************************

Private Function pBlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long

    On Error GoTo PROC_ERR_pBlendColors

Dim lIdx    As Long
Dim lRed    As Long
Dim lGrn    As Long
Dim lBlu    As Long
Dim fRedStp As Single
Dim fGrnStp As Single
Dim fBluStp As Single

    'Stop possible error
    If lSteps < 2 Then
        lSteps = 2
    End If

    'Extract Red, Blue and Green values from the start and end colors.
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000

    'Find the amount of change for each color element per color change.
    fRedStp = pDiv(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
    fGrnStp = pDiv(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
    fBluStp = pDiv(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))

    'Create the colors
    ReDim laRetColors(lSteps - 1)
    laRetColors(0) = lColor1            'First Color
    laRetColors(lSteps - 1) = lColor2   'Last Color
    For lIdx = 1 To lSteps - 2          'All Colors between
        laRetColors(lIdx) = CLng(lRed + (fRedStp * CSng(lIdx))) + _
                            (CLng(lGrn + (fGrnStp * CSng(lIdx))) * &H100&) + _
                            (CLng(lBlu + (fBluStp * CSng(lIdx))) * &H10000)
    Next

    'Return number of colors in array
    pBlendColors = lSteps

PROC_EXIT:

Exit Function

PROC_ERR_pBlendColors:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pBlendColors", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Function

'********************************************************************
'* Name: pCreateMemDC
'* Description: Create memory dc.
'********************************************************************

Private Sub pCreateMemDC(ByVal lW As Long, ByVal lH As Long)

    On Error GoTo PROC_ERR_pCreateMemDC

Dim lHdc As Long

    ' Destroy memory dc
    pDestroyMemDC

    ' Create a new dc based on desktop
    lHdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    ' Create compatible dc with desktop one
    m_lhDC = CreateCompatibleDC(lHdc)
    ' Create a new bitmap on which we will draw
    m_hBmp = CreateCompatibleBitmap(lHdc, lW, lH)
    ' Select bitmap into new dc
    m_hBmpOld = SelectObject(m_lhDC, m_hBmp)
    ' If there was an error
    If m_hBmpOld = 0 Then
        pDestroyMemDC
    Else
        ' Save dc width and height
        m_lWidth = lW
        m_lHeight = lH
    End If
    ' Set background mode to transparent
     SetBkMode m_lhDC, 1
    ' Delete temp dc
    DeleteDC lHdc

PROC_EXIT:

Exit Sub

PROC_ERR_pCreateMemDC:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pCreateMemDC", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

'********************************************************************
'* Name: pDestroyMemDC
'* Description: Delete all GDI object to free memory and to prevent
'* memory leaks.
'********************************************************************

Private Sub pDestroyMemDC()

    On Error GoTo PROC_ERR_pDestroyMemDC

    If m_hBmpOld <> 0 Then
        SelectObject m_lhDC, m_hBmpOld
        m_hBmpOld = 0
    End If
    If m_hBmp <> 0 Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    If m_lOrginalFont <> 0 Then
        SelectObject m_lhDC, m_lOrginalFont
        m_lOrginalFont = 0
    End If
    If m_lMemoryFont <> 0 Then
        DeleteObject m_lMemoryFont
        m_lMemoryFont = 0
    End If
    m_lWidth = 0
    m_lHeight = 0
    If m_lhDC <> 0 Then
        DeleteDC m_lhDC
        m_lhDC = 0
    End If

PROC_EXIT:

Exit Sub

PROC_ERR_pDestroyMemDC:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pDestroyMemDC", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

'********************************************************************
'* Name: pDiv
'* Description: Divides dNumer by dDenom if dDenom <> 0. Eliminates
'* 'Division By Zero' error.
'* Credits: Kath-Rock
'********************************************************************

Private Function pDiv(ByVal dNumer As Double, ByVal dDenom As Double) As Double

    On Error GoTo PROC_ERR_pDiv

    If dDenom <> 0 Then
        pDiv = dNumer / dDenom
    End If

PROC_EXIT:

Exit Function

PROC_ERR_pDiv:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pDiv", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Function

'********************************************************************
'* Name: IsWinNT
'* Description: Return true if windows platform is NT.
'********************************************************************

Private Function pIsWinNT() As Boolean

Dim lpOS As OSVERSIONINFO

    lpOS.dwOSVersionInfoSize = Len(lpOS)
    GetVersionEx lpOS
    pIsWinNT = (lpOS.dwPlatformID = 2)

End Function

'********************************************************************
'* Name: pOsVersion
'* Description: Returns OS version number.
'********************************************************************

Private Property Get pOsVersion() As Long

Dim tOVI As OSVERSIONINFO

    tOVI.dwOSVersionInfoSize = Len(tOVI)
    If GetVersionEx(tOVI) Then
        pOsVersion = tOVI.dwMajorVersion * 256 + tOVI.dwMinorVersion
    End If

End Property

'********************************************************************
'* Name: pRotatePoint
'* Description:
'* Credits: Kath-Rock
'********************************************************************

Private Sub pRotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Single)

    On Error GoTo PROC_ERR_pRotatePoint

Dim fDX         As Single
Dim fDY         As Single
Dim fRadians    As Single

    fRadians = fDegrees * RADS
    fDX = uRotatePt.x - uAxisPt.x
    fDY = uRotatePt.y - uAxisPt.y
    uRotatePt.x = uAxisPt.x + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
    uRotatePt.y = uAxisPt.y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))

PROC_EXIT:

Exit Sub

PROC_ERR_pRotatePoint:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.pRotatePoint", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

Private Sub pvTransBlt( _
                       ByVal hdcDest As Long, _
                       ByVal xDest As Long, _
                       ByVal yDest As Long, _
                       ByVal nWidth As Long, _
                       ByVal nHeight As Long, _
                       ByVal hdcSrc As Long, _
                       Optional ByVal xSrc As Long = 0, _
                       Optional ByVal ySrc As Long = 0, _
                       Optional ByVal clrMask As OLE_COLOR = vbMagenta, _
                       Optional ByVal hPal As Long = 0)

Dim hdcMask             As Long ' hDC of the created mask image
Dim hdcColor            As Long ' hDC of the created color image
Dim hBmMask             As Long ' Bitmap handle to the mask image
Dim hbmColor            As Long ' Bitmap handle to the color image
Dim hbmColorOld         As Long
Dim hbmMaskOld          As Long
Dim hpalOld             As Long
Dim hdcScreen           As Long
Dim hdcScnBuffer        As Long ' Buffer to do all work on
Dim hbmScnBuffer        As Long
Dim hbmScnBufferOld     As Long
Dim hPalBufferOld       As Long
Dim lMaskColor          As Long
Dim hpalHalftone        As Long

    hdcScreen = GetDC(0&)
    ' Validate palette
    If hPal = 0 Then
        hpalHalftone = CreateHalftonePalette(hdcScreen)
        hPal = hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    lMaskColor = lMaskColor And &HFFFFFF
    ' Create a color bitmap to server as a copy of the destination
    ' Do all work on this bitmap and then copy it back over the destination
    ' when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    ' Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    ' Copy the destination to the screen buffer
    APIBitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy
    ' Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    ' hdcSrc, because this will create a DIB section if the original bitmap
    ' is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    ' Now create a monochrome bitmap for the mask
    hBmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    ' First, blt the source bitmap onto the cover.  We do this first
    ' and then use it instead of the source bitmap
    ' because the source bitmap may be
    ' a DIB section, which behaves differently than a bitmap.
    ' (Specifically, copying from a DIB section to a monochrome bitmap
    ' does a nearest-color selection rather than painting based on the
    ' backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hpalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    ' In case hdcSrc contains a monochrome bitmap, we must set the destination
    ' foreground/background colors according to those currently set in hdcSrc
    ' (because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    APIBitBlt hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy
    ' Paint the mask.  What we want is white at the transparent color
    ' from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hBmMask)
    ' When ApiBitBlt'ing from color to monochrome, Windows sets to 1
    ' all pixels that match the background color of the source DC.  All
    ' other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    APIBitBlt hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy
    ' Paint the rest of the cover bitmap.

    ' What we want here is black at the transparent color, and
    ' the original colors everywhere else.  To do this, we first
    ' paint the original onto the cover (which we already did), then we
    ' AND the inverse of the mask onto that using the DSna ternary raster
    ' operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    ' Operation Codes", "Ternary Raster Operations", or search in MSDN
    ' for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".

    ' When ApiBitBlt'ing from monochrome to color, Windows transforms all white
    ' bits (1) to the background color of the destination hDC.  All black (0)
    ' bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    APIBitBlt hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, &H220326
    ' Paint the Mask to the Screen buffer
    APIBitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd
    ' Paint the Color to the Screen buffer
    APIBitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint
    ' Copy the screen buffer to the screen
    APIBitBlt hdcDest, xDest, yDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy
    ' All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hpalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, 0
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer
    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
    If hpalHalftone <> 0 Then
        DeleteObject hpalHalftone
    End If

End Sub

'********************************************************************
'* Name: SetPixel
'* Description: Set current pixel color.
'********************************************************************

Public Sub SetPixel(ByVal x As Long, ByVal y As Long, ByVal lColor As Long)

    On Error GoTo PROC_ERR_SetPixel

    If Not IsCreated() Then
        Exit Sub
    End If

    ' Set pixel color
    APISetPixel m_lhDC, x, y, lColor

PROC_EXIT:

Exit Sub

PROC_ERR_SetPixel:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.SetPixel", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Sub

'********************************************************************
'* Name: ShiftColor
'* Description: Returns shifted color(added or removed certain color)
'* from specified color.
'********************************************************************

Public Function ShiftColor(ByVal lColor As Long, ByVal lValue As Long, Optional bIsXP As Boolean = False) As Long

Dim lRed As Long
Dim lBlue As Long
Dim lGreen As Long

    lColor = TranslateColor(lColor)

    If bIsXP Then
        lBlue = ((lColor \ &H10000) Mod &H100)
        lBlue = lBlue + ((lBlue * lValue) \ &HC0)
    Else
        lBlue = ((lColor \ &H10000) Mod &H100) + lValue
    End If
    lGreen = ((lColor \ &H100) Mod &H100) + lValue
    lRed = (lColor And &HFF) + lValue

    ' Keep color values in range of 0-255
    If lRed < 0 Then
        lRed = 0
    ElseIf lRed > 255 Then
        lRed = 255
    End If
    If lGreen < 0 Then
        lGreen = 0
    ElseIf lGreen > 255 Then
        lGreen = 255
    End If
    If lBlue < 0 Then
        lBlue = 0
    ElseIf lBlue > 255 Then
        lBlue = 255
    End If

    ' Return shifted color
    ShiftColor = RGB(lRed, lGreen, lBlue)

End Function

'********************************************************************
'* Name: TextHeight
'* Description: Return provided text height using current font.
'********************************************************************

Public Function TextHeight(ByVal sText As String) As Long
     
    On Error GoTo PROC_ERR_TextHeight
    
    If Len(sText) = 0 Then Exit Function

    If Not IsCreated() Then
        Exit Function
    End If

Dim tSA As SIZEAPI
    If Not pIsWinNT Then
       GetTextExtentPointA m_lhDC, sText, Len(sText), tSA
    Else
       GetTextExtentPointW m_lhDC, StrPtr(sText), Len(sText), tSA
    End If
    TextHeight = tSA.cY

PROC_EXIT:

Exit Function

PROC_ERR_TextHeight:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.TextHeight", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Function

'********************************************************************
'* Name: TextWidth
'* Description: Return provided text width using current font.
'********************************************************************

Public Function TextWidth(ByVal sText As String) As Long

    On Error GoTo PROC_ERR_TextWidth
    
    If Len(sText) = 0 Then Exit Function

    If Not IsCreated() Then
        Exit Function
    End If

Dim tSA As SIZEAPI
    If Not pIsWinNT Then
       GetTextExtentPointA m_lhDC, sText, Len(sText), tSA
    Else
       GetTextExtentPointW m_lhDC, StrPtr(sText), Len(sText), tSA
    End If
    TextWidth = tSA.cX

PROC_EXIT:

Exit Function

PROC_ERR_TextWidth:
    Err.Raise ERRBASE, App.EXEName & ".CMemoryDC.TextWidth", "CMemoryDC component failure!" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "On line: " & Erl
    Resume PROC_EXIT

End Function

Public Function ShiftColorEx(Color As Long, PercentInDecimal As Single) As Long

'****************************************************************************
'* This routine shifts a color value specified by PercentInDecimal          *
'* Function inspired from DCbutton                                          *
'* All Credits goes to Noel Dacara                                          *
'* A Littlebit modified by me                                               *
'****************************************************************************

   Dim r                As Long
   Dim g                As Long
   Dim B                As Long

   '  Add or remove a certain color quantity by how many percent.

   r = Color And 255
   g = (Color \ 256) And 255
   B = (Color \ 65536) And 255

   r = r + PercentInDecimal * 255       ' Percent should already
   g = g + PercentInDecimal * 255       ' be translated.
   B = B + PercentInDecimal * 255       ' Ex. 50% -> 50 / 100 = 0.5

   '  When overflow occurs, ....
   If (PercentInDecimal > 0) Then       ' RGB values must be between 0-255 only
      If (r > 255) Then r = 255
      If (g > 255) Then g = 255
      If (B > 255) Then B = 255
   Else
      If (r < 0) Then r = 0
      If (g < 0) Then g = 0
      If (B < 0) Then B = 0
   End If

   ShiftColorEx = r + 256& * g + 65536 * B ' Return shifted color value

End Function

'********************************************************************
'* Name: TranslateColor
'* Description: Translates color from OLE_COLOR type to system color.
'********************************************************************

Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long

    On Error Resume Next

        'If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = -1':( --> replaced by:
        If OleTranslateColor(clr, hPal, TranslateColor) Then
            TranslateColor = -1
        End If

End Function

Public Property Let Width(ByVal lW As Long)

    'If lW > m_lWidth Then pCreateMemDC lW, m_lHeight':( --> replaced by:
    If lW > m_lWidth Then
        pCreateMemDC lW, m_lHeight
    End If

End Property

'********************************************************************
'* Name: Width
'* Description: Get\Set memory dc width.
'********************************************************************

Public Property Get Width() As Long

    Width = m_lWidth

End Property




