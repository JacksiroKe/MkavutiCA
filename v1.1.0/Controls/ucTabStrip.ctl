VERSION 5.00
Begin VB.UserControl ucTabStrip 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucTabStrip.ctx":0000
   Begin VB.Shape shpHover 
      BorderColor     =   &H80000010&
      FillColor       =   &H8000000F&
      Height          =   315
      Left            =   960
      Top             =   30
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line SepLineShadow 
      BorderColor     =   &H80000014&
      Index           =   0
      Visible         =   0   'False
      X1              =   57
      X2              =   57
      Y1              =   2
      Y2              =   22
   End
   Begin VB.Line SepLine 
      BorderColor     =   &H80000010&
      Index           =   0
      Visible         =   0   'False
      X1              =   56
      X2              =   56
      Y1              =   2
      Y2              =   22
   End
   Begin VB.Label lblTabButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tab 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   600
   End
End
Attribute VB_Name = "ucTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'
'   Build Date & Time: 6/28/2006 10:57:40 PM
Const Major As Long = 1
Const Minor As Long = 0
Const Revision As Long = 142
Const DateTime As String = "6/28/2006 10:57:40 PM"

Private Type POINT
    X As Long
    Y As Long
End Type

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Enum tsAppearanceEnum
    [tsFlat] = &H0
    [ts3D] = &H1
End Enum

Public Enum tsBackStyleEnum
    [tsTransparent] = &H0
    [tsOpaque] = &H1
End Enum

Public Enum tsBorderStyleEnum
    [tsNone] = &H0
    [tsFixedSingle] = &H1
End Enum

Private Enum tsGradientDirectionEnum
    [tsNWSE] = &H0                              'Diagonal Gradient from Upper Left to Lower Right
    [tsSWNE] = &H1                              'Diagonal Gradient from Lower Left to Upper Right
End Enum

Public Enum tsGradientStyleEnum
    [tsNoGradient] = &H0                        'No Gradient, Use BackColor instead
    [tsCircular] = &H1                          'Circular Gradient with Center @ ScaleWidth/3, ScaleHeight/3
    [tsDiagonalNWSE] = &H2                      'Diagonal Gradient from Upper Left to Lower Right
    [tsDiagonalSWNE] = &H3                      'Diagonal Gradient from Lower Left to Upper Right
    [tsHorizontal] = &H4                        'Horizontal Gradient
    [tsRectangular] = &H5                       'Rectangular Gradient
    [tsVertical] = &H6                          'Vertical Gradient
End Enum

'   Note that bCancel is passed by Reference in below event. This event is called just before a
'   tab is being switched, we can prevent tab switch by making bCancel as true
'   If we Set bCancel in the BeforeTabSwitch following event will not occur.
Public Event BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
Public Event Click()
Public Event DblClick()
Public Event EnterFocus()
Public Event ExitFocus()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event TabDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event TabHover(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event TabUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event TabSwitch(iLastActiveTab As Integer)

Private m_ActiveForeColor As OLE_COLOR          'ActiveForeColor property
Private m_ActiveTab As Integer                  'ActiveTab property
Private m_Appearance As tsAppearanceEnum        'Appearance property
Private m_BackColor As OLE_COLOR                'Backcolor property
Private m_bCancel As Boolean                    'Cancel Flag for Tab Change
Private m_BorderStyle As tsBorderStyleEnum      'BorderStyle property
Private m_Caption() As String                   'Caption property
Private m_CaptionAlign As Long                  'CaptionAlign property
Private m_Font As StdFont                       'TabFont property
Private m_GradientEnd As OLE_COLOR              'GradientEnd property
Private m_GradientStart As OLE_COLOR            'GradientStart property
Private m_GradientStyle As tsGradientStyleEnum  'GradientStyle property
Private m_Height As Long                        'ScaleHeight
Private m_HoverColor As OLE_COLOR               'HoverColor property
Private m_InActiveForeColor As OLE_COLOR        'InActiveForeColor property
Private m_LastActiveTab As Integer              'Last Active Tab
Private m_Separators As Boolean                 'Separators property
Private m_SeparatorColor As OLE_COLOR           'SeparatorColor property
Private m_TabBackColor As OLE_COLOR             'TabBackColor property
Private m_TabBackStyle As tsBackStyleEnum       'Backstyle property
Private m_TabCount As Long                      'TabCount property
Private m_TabHeight As Long                     'TabHeight property
Private m_TabStyle As tsAppearanceEnum          'TabStyle property
Private m_Width As Long                         'ScaleWidth
Private bInitGradient As Boolean                'Initialize Gradient flag
Private bInternal As Boolean                    'Internal Change flag

Public Property Get ActiveForeColor() As OLE_COLOR
    '   The ActiveForeColor property
    ActiveForeColor = m_ActiveForeColor
End Property

Public Property Let ActiveForeColor(ByVal NewValue As OLE_COLOR)
    m_ActiveForeColor = NewValue
    Call BuildTabs
    Call Refresh(m_ActiveTab)
    PropertyChanged "ActiveForeColor"
End Property

Public Property Get ActiveTab() As Long
    ' The ActiveTab property
    ActiveTab = m_ActiveTab
End Property

Public Property Let ActiveTab(ByVal lNewValue As Long)
    If (lNewValue < 0) Or (lNewValue >= m_TabCount) Then
        If (lNewValue < 0) Then lNewValue = m_TabCount
        If (lNewValue >= m_TabCount) Then lNewValue = 0
    End If
    '   If already we are on the same tab (this is important or else all
    '   the contained controls for active tab will be moved to -75000 and so...
    If lNewValue = m_ActiveTab Then Exit Property
    m_bCancel = False
    '   Raise event and confirm that the user want to allow the tab switch
    RaiseEvent BeforeTabSwitch(lNewValue, m_bCancel)
    '   If user set the cancel flag in the BeforeTabSwitch event
    If m_bCancel Then Exit Property
    '   Show/Hide Controls for active tab
    Call HandleContainedControls(lNewValue)
    '   Store current tab in last active tab
    m_LastActiveTab = m_ActiveTab
    '   Now set the New Current Tab
    m_ActiveTab = lNewValue
    PropertyChanged "ActiveTab"
    '   Redraw
    If Not bInternal Then
        Call Refresh(m_ActiveTab)
    End If
    RaiseEvent TabSwitch(m_LastActiveTab)
End Property

Private Sub APILine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lColor As Long)
    '   Use the API LineTo for Fast Drawing
    On Error GoTo APILine_Error
    
    Dim Pt As POINT
    Dim hPen As Long, hPenOld As Long
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(UserControl.hDC, hPen)
    MoveToEx UserControl.hDC, X1, Y1, Pt
    LineTo UserControl.hDC, X2, Y2
    SelectObject UserControl.hDC, hPenOld
    DeleteObject hPen
    Exit Sub
    
APILine_Error:
End Sub

Private Function APIRectangle(ByVal X As Long, ByVal Y As Long, ByVal W As Long, _
    ByVal H As Long, Optional ByVal lColor As OLE_COLOR = -1) As Long
    
    '   Use the API Rectangle for Fast Drawing
    On Error GoTo APIRectangle_Error
    
    Dim hPen As Long, hPenOld As Long
    Dim Pt As POINT
    
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(hDC, hPen)
    Rectangle UserControl.hDC, X, Y, W, H
    SelectObject UserControl.hDC, hPenOld
    DeleteObject hPen
    Exit Function
    
APIRectangle_Error:
End Function

Public Property Get Appareance() As tsAppearanceEnum
    '   Appearance (Flat, 3D)
    Appareance = UserControl.Appearance
End Property

Public Property Let Appareance(ByVal New_Appearance As tsAppearanceEnum)
    m_BackColor = UserControl.BackColor
    UserControl.Appearance = New_Appearance
    m_Appearance = New_Appearance
    '   Make sure to set the Color back as native controls
    '   have a bad habit of changing this on us...
    UserControl.BackColor = m_BackColor
    '   Set the Gradient Drawing Flag
    bInitGradient = False
    '   Repaint things
    Call Refresh(m_ActiveTab)
    PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
    '   BackColor (Tab Body)
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get BorderStyle() As tsBorderStyleEnum
    ' The BorderStyle property
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As tsBorderStyleEnum)
    m_BorderStyle = NewValue
    UserControl.BorderStyle = m_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub BuildTabs()
    Dim lLeft As Long
    Dim lTop As Long
    Dim lCount As Long
    Dim lSepCount As Long
    Dim i As Long
    Dim Diff As Long
    Dim bNewRow As Boolean
    Dim iRow As Long
    
    '   This Sub Builds/Destroys the Tabs and Separators depending on the m_TabCount
    
    With UserControl
        '   Get the current count
        '   Note: This count does not reflect the Zero'th Index
        lCount = lblTabButton.Count - 1
        '   More Tabs....
        If m_TabCount > lCount Then
            Diff = m_TabCount - lCount - 1
            For i = 0 To Diff
                If i < Diff Then
                    Load .lblTabButton(lCount + 1)
                    lCount = .lblTabButton.Count - 1
                    ReDim Preserve m_Caption(0 To lCount)
                    With .lblTabButton(lCount)
                        '   Set the visual styles to compute
                        '   the size of the label when the border
                        '   and font are in the "selected" state.
                        .Appearance = m_TabStyle
                        .BackStyle = m_TabBackStyle
                        .BorderStyle = tsFixedSingle
                        '   Give it a default caption
                        .Caption = "Tab " & lCount
                        '   Store this for refreshing
                        m_Caption(lCount) = .Caption
                        Set .Font = m_Font
                        With .Font
                            .Bold = True
                        End With
                        .BackColor = m_TabBackColor
                        .ForeColor = m_InActiveForeColor
                        .Height = TabHeight
                        .Width = TextWidth(.Caption) + 15
                        '   Now set the font for display
                        .Font.Bold = False
                        .BorderStyle = tsNone
                        '   Make it visible
                        .Visible = True
                        '   We will move the controls in a central location....
                        '   namely, MoveControls Method since they are all stacked
                        '   up at Left = 0, Top = 0
                    End With
                End If
                If (m_TabCount >= 1) And (i < Diff) Then
                    '   Now load the separator lines
                    lSepCount = SepLine.Count - 1
                    Load .SepLine(lSepCount + 1)
                    Load .SepLineShadow(lSepCount + 1)
                    lSepCount = .SepLine.Count - 1
                    With .SepLine(lSepCount)
                        .Visible = m_Separators
                    End With
                    With .SepLineShadow(lSepCount)
                        .Visible = m_Separators
                    End With
                End If
            Next i
        ElseIf m_TabCount <= lCount Then
            '   Ah, we are removeing things, so do them one at a time
            Diff = (lCount - m_TabCount) + 1
            '   Remove these from the end...so get the count
            lCount = .lblTabButton.Count - 1
            lSepCount = .SepLine.Count - 1
            For i = 1 To Diff
                Unload .lblTabButton(lCount)
                '   Get the current count
                lCount = lblTabButton.Count - 1
            Next i
            '   Eliminate the Separators as well
            For i = 1 To Diff
                Unload .SepLine(lSepCount)
                Unload .SepLineShadow(lSepCount)
                '   Get the current count
                lSepCount = .SepLine.Count - 1
                ReDim Preserve m_Caption(0 To lSepCount)
            Next i
            
        End If
        '   Now move the controls
        Call MoveControls(0)
    End With
End Sub

Public Property Get CaptionAlign() As AlignmentConstants
    '   The Caption property
    CaptionAlign = lblTabButton(0).Alignment
End Property

Public Property Let CaptionAlign(ByVal NewValue As AlignmentConstants)
    Dim i As Long
    m_CaptionAlign = NewValue
    For i = 0 To lblTabButton.Count - 1
        lblTabButton(i).Alignment = NewValue
    Next i
End Property

Public Property Get Caption(ByVal Index As Integer) As String
    '   The Caption property
    Caption = lblTabButton(Index).Caption
End Property

Public Property Let Caption(ByVal Index As Integer, ByVal NewValue As String)
    m_Caption(Index) = NewValue
    lblTabButton(Index).Caption = NewValue
    lblTabButton(Index).Width = TextWidth(NewValue) + 15
    Call MoveControls(Index)
End Property

Private Sub CheckAccelerator(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Long
    Dim j As Long
    
    '   This routine checks for Accelerator keys and moves the active
    '   Tab to the current one...
    With UserControl
        For i = 0 To .lblTabButton.Count - 1
            Debug.Print Chr$(38) & Chr$((KeyCode + 32))
            If InStr(1, .lblTabButton(i).Caption, Chr$(38) & Chr$(KeyCode + 32), vbTextCompare) <> 0 Then
                ActiveTab = i
            End If
        Next i
    End With
End Sub

Private Sub DrawCGradient(ByVal StartColor As Long, ByVal EndColor As Long, _
    Optional ByVal numSteps As Integer = 256, Optional ByVal XCenter As Single = -1, _
    Optional ByVal YCenter As Single = -1)
    
    '   Draw a Circular Gradient
    Dim StartRed As Integer, StartGreen As Integer, StartBlue As Integer
    Dim DeltaRed As Integer, DeltaGreen As Integer, DeltaBlue As Integer
    Dim stp As Long, hPen As Long, hPenOld As Long, lColor As Long
    Dim X As Long, Y As Long, X2 As Long, Y2 As Long

    With UserControl
        ' Evaluate the coordinates off the center if omitted.
        If XCenter = -1 And YCenter = -1 Then
            XCenter = .ScaleWidth / 3
            YCenter = .ScaleHeight / 3
        End If
                
        ' Split the start color into its RGB components
        StartRed = StartColor And &HFF
        StartGreen = (StartColor And &HFF00&) \ 256
        StartBlue = (StartColor And &HFF0000) \ 65536
        ' Split the end color into its RGB components
        DeltaRed = (EndColor And &HFF&) - StartRed
        DeltaGreen = (EndColor And &HFF00&) \ 256 - StartGreen
        DeltaBlue = (EndColor And &HFF0000) \ 65536 - StartBlue

        ' Draw all circles, going from the outside in.
        For stp = 0 To numSteps - 1
            lColor = RGB(StartRed + (DeltaRed * stp) \ numSteps, _
                StartGreen + (DeltaGreen * stp) \ numSteps, _
                StartBlue + (DeltaBlue * stp) \ numSteps)
            X = XCenter - numSteps + stp
            Y = YCenter - numSteps + stp
            X2 = XCenter + numSteps - stp
            Y2 = YCenter + numSteps - stp
            hPen = CreatePen(0, 2, lColor)
            hPenOld = SelectObject(UserControl.hDC, hPen)
            Ellipse .hDC, X, Y, X2, Y2
            SelectObject UserControl.hDC, hPenOld
            DeleteObject hPen
        Next
        
    End With
End Sub

Private Sub DrawDGradient(ByVal lStartColor As Long, ByVal lEndColor As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Direction As tsGradientDirectionEnum)
    
    '   Draw a Diagonal Gradient in the current HDC
    On Error GoTo DrawDGradient_Error
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long, lColor As Long, hPen As Long, hPenOld As Long
    
    lh = Y2 - Y
    lw = X2 - X
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    If lh > lw Then
        dR = (sR - eR) / (lh)
        dG = (sG - eG) / (lh)
        dB = (sB - eB) / (lh)
        For ni = 0 To lh + 1
            lColor = RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
            If Direction = tsNWSE Then
                '   NWSE (Move Only the UL corner towards the LR)
                Call APIRectangle(X + ni, Y + ni, X2, Y2, lColor)
            Else
                '   SWNE (Move Only the LL corner towards the UR)
                Call APIRectangle(X + ni, Y, X2, Y2 - ni, lColor)
            End If
        Next 'ni
    Else
        dR = (sR - eR) / (lw)
        dG = (sG - eG) / (lw)
        dB = (sB - eB) / (lw)
        For ni = 0 To lw + 1
            lColor = RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
            If Direction = tsNWSE Then
                '   NWSE (Move Only the UL corner towards the LR)
                Call APIRectangle(X + ni, Y + ni, X2, Y2, lColor)
            Else
                '   SWNE (Move Only the LL corner towards the UR)
                Call APIRectangle(X + ni, Y, X2, Y2 - ni, lColor)
            End If
        Next 'ni
    End If
    Exit Sub
    
DrawDGradient_Error:
End Sub

Private Sub DrawHGradient(ByVal lStartColor As Long, ByVal lEndColor As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    
    '   Draw a Horizontal Gradient in the current HDC
    On Error GoTo DrawHGradient_Error
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long
    lh = Y2 - Y
    lw = X2 - X
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / lw
    dG = (sG - eG) / lw
    dB = (sB - eB) / lw
    
    For ni = 0 To lw
        APILine X + ni, Y, X + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next 'ni
    
    Exit Sub
    
DrawHGradient_Error:
End Sub

Private Sub DrawRGradient(ByVal lStartColor As Long, ByVal lEndColor As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    
    '   Draw a Rectangular Gradient in the current HDC
    
    On Error Resume Next
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long, lColor As Long, hPen As Long, hPenOld As Long
    
    lh = Y2 - Y
    lw = X2 - X
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    If lh > lw Then
        dR = (sR - eR) / (lh / 3)
        dG = (sG - eG) / (lh / 3)
        dB = (sB - eB) / (lh / 3)
        For ni = 0 To (lh / 3) + 1
            lColor = RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
            If ni = (lh / 3) + 1 Then
                Call APIRectangle(X + ni - 1, Y + ni - 1, X2 - ni, Y2 - ni, lColor)
            Else
                Call APIRectangle(X + ni, Y + ni, X2 - ni, Y2 - ni, lColor)
            End If
        Next 'ni
    Else
        dR = (sR - eR) / (lw / 3)
        dG = (sG - eG) / (lw / 3)
        dB = (sB - eB) / (lw / 3)
        For ni = 0 To (lw / 3) + 1
            lColor = RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
            If ni = (lw / 3) + 1 Then
                Call APIRectangle(X + ni - 1, Y + ni - 1, X2 - ni, Y2 - ni, lColor)
            Else
                Call APIRectangle(X + ni, Y + ni, X2 - ni, Y2 - ni, lColor)
            End If
        Next 'ni
        
    End If
    Exit Sub
    
End Sub

Private Sub DrawVGradient(ByVal lStartColor As Long, ByVal lEndColor As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal X2 As Long, ByVal Y2 As Long)
    
    '   Draw a Vertical Gradient in the current HDC
    On Error GoTo DrawVGradient_Error
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    
    For ni = 0 To Y2
        APILine X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next 'ni
    
    Exit Sub
    
DrawVGradient_Error:
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    Call EnableControls(NewValue)
    PropertyChanged "Enabled"
End Property

Private Sub EnableControls(ByVal NewValue As Boolean)
    Dim i As Long
    Dim lCount As Long
    
    '   Quietly handle error for controls that don't
    '   have Enabled properties
    On Error Resume Next
    With UserControl
        lCount = .ContainedControls.Count
        If lCount >= 0 Then
            For i = 0 To lCount
                '   Set the controls Disabled/Enabled state
                .ContainedControls(i).Enabled = NewValue
            Next
        End If
    End With
End Sub

Public Property Get Font() As StdFont
    ' The Font property
    Set Font = m_Font
End Property

Public Property Set Font(ByVal NewValue As StdFont)
    Set m_Font = NewValue
    '   Flag that this is an internal change
    bInternal = True
    '   Refresh the control and its font settings
    Refresh
    '   Turn off the flag to prevent resetting things
    '   on the next refresh...
    bInternal = False
    PropertyChanged "Font"
End Property

Public Property Get GradientEnd() As OLE_COLOR
    GradientEnd = m_GradientEnd
End Property

Public Property Let GradientEnd(ByVal lNewColor As OLE_COLOR)
    m_GradientEnd = lNewColor
    '   Allow the Gradient to be redrawn
    bInitGradient = False
    Call Refresh(m_ActiveTab)
    PropertyChanged "GradientEnd"
End Property

Public Property Get GradientStart() As OLE_COLOR
    GradientStart = m_GradientStart
End Property

Public Property Let GradientStart(ByVal lNewColor As OLE_COLOR)
    m_GradientStart = lNewColor
    '   Allow the Gradient to be redrawn
    bInitGradient = False
    Call Refresh(m_ActiveTab)
    PropertyChanged "GradientStart"
End Property

Public Property Get GradientStyle() As tsGradientStyleEnum
    GradientStyle = m_GradientStyle
End Property

Public Property Let GradientStyle(ByVal NewStyle As tsGradientStyleEnum)
    m_GradientStyle = NewStyle
    '   Allow the Gradient to be redrawn
    bInitGradient = False
    Call Refresh(m_ActiveTab)
    PropertyChanged "GradientStyle"
End Property

Private Sub HandleContainedControls(ByVal New_ActiveTab As Long)
    ' VERY IMPORTANT FUNCTION:
    '   Handles the appearing and disappearing of controls for the current
    '   tab and last active tab
    '
    '   This routine replaces the original routine implemented with Collections
    '   and is based on the PCS article by Evan Todder:
    '   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=57642&lngWId=1
    '
    '   NOTE:
    '   Unfortunatly, the above article was removed by the Author from PCS, so the link
    '   is not active; however I felt it was important to give credit for the cleaver
    '   idea none the less.
    '
    Dim Ctl As Control
    Dim MoveVal As Long
 
    On Error Resume Next
    '   The difference between what was the active
    '   Tab and the newly set activetab
    MoveVal = (New_ActiveTab - m_ActiveTab)
    '   The code below has been changed to permit >45 Tabs
    '   at values above this the Left value overflows the
    '   Left property which and the alignment is lost ;-(
    'MoveVal = (MoveVal * 10000)
    MoveVal = (MoveVal * (Width + 100))

    '   This is what creates the illusion of
    '   Changing the Tab of a tab control
    For Each Ctl In UserControl.ContainedControls
         Ctl.left = (Ctl.left + MoveVal)
    Next Ctl

End Sub

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get HoverColor() As OLE_COLOR
    '   The HoverColor property for the Shape Control
    HoverColor = m_HoverColor
End Property

Public Property Let HoverColor(ByVal NewValue As OLE_COLOR)
    m_HoverColor = NewValue
    UserControl.shpHover.BorderColor = m_HoverColor
    PropertyChanged "HoverColor"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get InActiveForeColor() As OLE_COLOR
    '   The InActiveForeColor property
    InActiveForeColor = m_InActiveForeColor
End Property

Public Property Let InActiveForeColor(ByVal NewValue As OLE_COLOR)
    m_InActiveForeColor = NewValue
    Call BuildTabs
    Call Refresh(m_ActiveTab)
    PropertyChanged "InActiveForeColor"
End Property

Private Sub lblTabButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Refresh(Index)
    RaiseEvent TabDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblTabButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With UserControl
        If .lblTabButton(Index).BorderStyle = tsNone Then
            .shpHover.Move .lblTabButton(Index).left - 1, .lblTabButton(Index).Top - 1, .lblTabButton(Index).Width + 3
            .shpHover.Visible = True
        Else
            .shpHover.Visible = False
        End If
    End With
    RaiseEvent TabHover(Index, Button, Shift, X, Y)
End Sub

Private Sub lblTabButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent TabUp(Index, Button, Shift, X, Y)
End Sub

Private Sub MoveControls(ByVal Index As Integer)
    Dim i As Long
    Dim j As Long
    Dim iRow As Long
    Dim lLeft As Long
    Dim bNewRow As Boolean
    
    '   This Sub moves the controls relative to one another. If the index is passed
    '   the index to Control.Count -1 are moved, else all controls are moved.
    
    With UserControl
        '   Make the first visible and aligned always as this sets up the reamining controls
        With .SepLine(0)
            .X1 = lblTabButton(0).left + lblTabButton(0).Width + 4
            .X2 = lblTabButton(0).left + lblTabButton(0).Width + 4
            .Y1 = lblTabButton(0).Top
            .Y2 = lblTabButton(0).Top + lblTabButton(0).Height
            If (m_TabCount > 1) And (m_Separators) Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        With .SepLineShadow(0)
            .X1 = SepLine(0).X1 + 1
            .X2 = SepLine(0).X2 + 1
            .Y1 = SepLine(0).Y1
            .Y2 = SepLine(0).Y2
            If (m_TabCount > 1) And (m_Separators) Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        '   Move the Buttons from the Index Onward
        For i = Index To .lblTabButton.Count - 1
            lLeft = 0
            '   Compute the left extent of the controls to see if they will fit on one line.
            For j = iRow To i
                If j = 0 Then
                    lLeft = lLeft + (lblTabButton(j).left + lblTabButton(j).Width + 8)
                Else
                    lLeft = lLeft + 8 + (lblTabButton(j).Width + 8)
                End If
            Next
            '   See if we are going to fit....if not then adjust its position
            '   so it is on the next row
            If (lLeft > ScaleWidth) Then
                '   Increment the row number
                iRow = i + 1
                '   Compute the row number...
                bNewRow = True
                '   Move the Tab According to Previous Row Offset
                .lblTabButton(i).Move lblTabButton(0).left, (lblTabButton(i - 1).Top + lblTabButton(i - 1).Height + 8) '* lScale
            Else
                bNewRow = False
                '   Only if this is not the first
                If i > 0 Then
                    '    Still on the same row, just move them over relative to its neighbor
                    .lblTabButton(i).Move lblTabButton(i - 1).left + lblTabButton(i - 1).Width + 8, lblTabButton(i - 1).Top
                Else
                    '   This is the default location....if we set the AutoSize = True, then
                    '   the controls Left position can change...so make sure we start in
                    '   the correct location....
                    .lblTabButton(i).Move 8, 2
                End If
            End If
            '   Make the controls non-AutoSize
            .lblTabButton(i).AutoSize = False
            '   Now move all of the separators as well....
            If m_Separators Then
                With .SepLine(i)
                    .X1 = lblTabButton(i).left + lblTabButton(i).Width + 4
                    .X2 = lblTabButton(i).left + lblTabButton(i).Width + 4
                    .Y1 = lblTabButton(i).Top
                    .Y2 = (lblTabButton(i).Top + lblTabButton(i).Height)
                End With
                With .SepLineShadow(i)
                    .X1 = lblTabButton(i).left + lblTabButton(i).Width + 5
                    .X2 = lblTabButton(i).left + lblTabButton(i).Width + 5
                    .Y1 = lblTabButton(i).Top
                    .Y2 = (lblTabButton(i).Top + lblTabButton(i).Height)
                End With
            End If
        Next
    End With
End Sub

Private Function OffsetColor(ByVal lColor As OLE_COLOR, ByVal lOffset As Long) As OLE_COLOR

    Dim lRed As OLE_COLOR
    Dim lGreen As OLE_COLOR
    Dim lBlue As OLE_COLOR
    Dim lR As OLE_COLOR
    Dim lg As OLE_COLOR
    Dim lb As OLE_COLOR
    
    '   Make sure the color is on the pallete
    lColor = TranslateColor(lColor)
    '   Now offset the color by splitting the color
    lR = (lColor And &HFF)
    lg = ((lColor And 65280) \ 256)
    lb = ((lColor) And 16711680) \ 65536
    lRed = (lOffset + lR)
    lGreen = (lOffset + lg)
    lBlue = (lOffset + lb)
    If lRed > 255 Then lRed = 255
    If lRed < 0 Then lRed = 0
    If lGreen > 255 Then lGreen = 255
    If lGreen < 0 Then lGreen = 0
    If lBlue > 255 Then lBlue = 255
    If lBlue < 0 Then lBlue = 0
    OffsetColor = RGB(lRed, lGreen, lBlue)
    
End Function

Public Sub Refresh(Optional ByVal Index As Integer)
    Dim i As Long
    Dim bAdjust As Boolean
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   This Sub refreshes the controls with the correct backcolor,, forecolor, borderstyle and font
    With UserControl
        For i = 0 To .lblTabButton.Count - 1
            With .lblTabButton(i)
                .Appearance = tsFlat
                .Alignment = m_CaptionAlign
                .BackStyle = m_TabBackStyle
                .BorderStyle = tsNone
                '   See if we have a passed caption, if not then
                '   just use the default caption (i.e. Tab 1,...Tab n)
                If Len(m_Caption(i)) > 0 Then
                    .Caption = m_Caption(i)
                End If
                Set .Font = m_Font
                With .Font
                    .Bold = False
                    .Size = m_Font.Size
                    .Underline = m_Font.Underline
                End With
                .ForeColor = m_InActiveForeColor
                .BackColor = m_TabBackColor
                '   Make sure the font changes will fit in the control
                If bInternal Then
                    .AutoSize = True
                    bAdjust = True
                Else
                    .Height = m_TabHeight
                End If
                .Visible = True
            End With
            '   Make sure the Separators Height = TabHeight
            With .SepLine(i)
                .Y2 = .Y1 + m_TabHeight
                .BorderColor = m_SeparatorColor
            End With
            With .SepLineShadow(i)
                .Y2 = .Y1 + m_TabHeight
                '.BorderColor = OffsetColor(m_SeparatorColor, &HF0)
                .BorderColor = &HFFFFFF
                '   Hide the shadow if Flat
                If m_TabStyle = tsFlat Then
                    .Visible = False
                Else
                    .Visible = True
                End If
            End With
        Next i
        '   Make sure the Hover Shape Height = TabHeight
        .shpHover.Height = lblTabButton(0).Height + 1
        '   New Active Tab
        If Index <= lblTabButton.Count - 1 Then
            With .lblTabButton(Index)
                .Appearance = m_TabStyle
                .BackStyle = m_TabBackStyle
                .BorderStyle = tsFixedSingle
                .Font.Bold = True
                .ForeColor = m_ActiveForeColor
                .BackColor = OffsetColor(m_TabBackColor, &HC) '&HD8E9EC ~ vbButtonFace
            End With
            '   Now set the Controls to match this change
            If Index <> m_ActiveTab Then
                bInternal = True
                ActiveTab = Index
                bInternal = False
            End If
        End If
        '   See if the Gradient has been built, if not then do it....
        '   This prevents unwanted painting of the control which when
        '   the control is large can be slow...
        If Not bInitGradient Then
            '   Speed things up by locking the window
            LockWindowUpdate .hWnd
            '   Clear the control surface
            .Cls
            '   Paint the correct gradient
            Select Case m_GradientStyle
                Case tsNoGradient
                    '   Do nothing....
                Case tsCircular
                    Call DrawCGradient(m_GradientStart, m_GradientEnd, .ScaleWidth)
                Case tsDiagonalNWSE
                    Call DrawDGradient(m_GradientEnd, m_GradientStart, 0, 0, .ScaleWidth, .ScaleHeight, tsNWSE)
                Case tsDiagonalSWNE
                    Call DrawDGradient(m_GradientEnd, m_GradientStart, 0, 0, .ScaleWidth, .ScaleHeight, tsSWNE)
                Case tsHorizontal
                    Call DrawHGradient(m_GradientStart, m_GradientEnd, 0, 0, .ScaleWidth, .ScaleHeight)
                Case tsRectangular
                    Call DrawRGradient(m_GradientStart, m_GradientEnd, 0, 0, .ScaleWidth, .ScaleHeight)
                Case tsVertical
                    Call DrawVGradient(m_GradientStart, m_GradientEnd, 0, 0, .ScaleWidth, .ScaleHeight)
            End Select
            '   Now unlock things for the update to take effect
            LockWindowUpdate 0&
            '   Make sure to mark that we have built the gradient
            '   to prevent unwanted drawing of the control....
            bInitGradient = True
        End If
        '   See if the font or size has changed....if so then move things
        If bAdjust Then
            '   Store the new TabHeight
            m_TabHeight = lblTabButton(0).Height
            Call MoveControls(0)
        End If
    End With
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get Separators() As Boolean
    ' The Separators property
    Separators = m_Separators
End Property

Public Property Let Separators(ByVal NewValue As Boolean)
    m_Separators = NewValue
    Call SeparatorsVisible(NewValue)
    PropertyChanged "Separators"
End Property

Public Property Get SeparatorColor() As OLE_COLOR
    ' The SeparatorColor property
    SeparatorColor = m_SeparatorColor
End Property

Public Property Let SeparatorColor(ByVal NewValue As OLE_COLOR)
    m_SeparatorColor = NewValue
    Call Refresh(m_ActiveTab)
    PropertyChanged "SeparatorColor"
End Property

Private Sub SeparatorsVisible(ByVal NewState As Boolean)
    Dim i As Long
    
    '   Hide/Show Separators on the control
    With UserControl
        For i = 0 To .SepLine.Count - 1
            .SepLine(i).Visible = NewState
            .SepLineShadow(i).Visible = NewState
        Next
        Call MoveControls(0)
        .Refresh
    End With
End Sub

Public Property Get TabBackColor() As OLE_COLOR
    '   The TabBackColor property
    TabBackColor = m_TabBackColor
End Property

Public Property Let TabBackColor(ByVal NewValue As OLE_COLOR)
    m_TabBackColor = NewValue
    Call BuildTabs
    Call Refresh(m_ActiveTab)
    PropertyChanged "TabBackColor"
End Property

Public Property Get TabBackStyle() As tsBackStyleEnum
    '   The TabBackStyle property
    TabBackStyle = m_TabBackStyle
End Property

Public Property Let TabBackStyle(ByVal NewValue As tsBackStyleEnum)
    m_TabBackStyle = NewValue
    Call Refresh(m_ActiveTab)
    PropertyChanged "TabBackStyle"
End Property

Public Property Get TabCount() As Long
    ' The TabCount property
    TabCount = m_TabCount
End Property

Public Property Let TabCount(ByVal NewValue As Long)
    If NewValue > 45 Then NewValue = 45
    m_TabCount = NewValue
    Call BuildTabs
    Call Refresh(m_ActiveTab)
    PropertyChanged "TabCount"
End Property

Public Property Get TabHeight() As Long
    ' The TabHeight property
    TabHeight = m_TabHeight
End Property

Public Property Let TabHeight(ByVal NewValue As Long)
    m_TabHeight = NewValue
    Call Refresh(m_ActiveTab)
    PropertyChanged "TabHeight"
End Property

Public Property Get TabStyle() As tsAppearanceEnum
    ' The TabStyle property
    TabStyle = m_TabStyle
End Property

Public Property Let TabStyle(ByVal NewValue As tsAppearanceEnum)
    m_TabStyle = NewValue
    Call Refresh(m_ActiveTab)
    PropertyChanged "TabStyle"
End Property

Private Function TranslateColor(ByVal lColor As Long) As Long
    '   System color code to long rgb
    On Error GoTo TranslateColor_Error
    
    If OleTranslateColor(lColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If
    Exit Function
    
TranslateColor_Error:
End Function

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
    RaiseEvent EnterFocus
End Sub

Private Sub UserControl_ExitFocus()
    RaiseEvent ExitFocus
End Sub

Private Sub UserControl_Initialize()
    With UserControl
        .shpHover.BorderColor = m_HoverColor
    End With
End Sub

Private Sub UserControl_InitProperties()
    With UserControl
        m_ActiveTab = 1
        CaptionAlign = vbCenter
        m_ActiveForeColor = &H80000012
        m_Appearance = tsFlat
        m_BackColor = &H8000000F
        m_BorderStyle = tsNone
        Set m_Font = UserControl.Parent.Font
        UserControl.ForeColor = &H80000011
        m_GradientEnd = &HFFFFFF
        m_GradientStart = &HFFC0C0
        m_Height = .ScaleHeight
        m_HoverColor = &H80000010
        m_InActiveForeColor = &H80000011
        m_Separators = True
        m_SeparatorColor = &H80000010
        m_TabBackColor = &H8000000F
        m_TabBackStyle = tsTransparent
        m_TabCount = 3
        m_TabHeight = 20
        m_TabStyle = ts3D
        m_Width = .ScaleWidth
    End With
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 4
            Call CheckAccelerator(KeyCode, Shift)
    End Select
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With UserControl
        .shpHover.Visible = False
    End With
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_ActiveTab = .ReadProperty("ActiveTab", 0)
        m_Appearance = .ReadProperty("Appearance", tsFlat)
        UserControl.Appearance = m_Appearance
        UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
        m_BackColor = UserControl.BackColor
        m_BorderStyle = .ReadProperty("BorderStyle", tsNone)
        UserControl.BorderStyle = m_BorderStyle
        m_CaptionAlign = UserControl.lblTabButton(0).Alignment
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set m_Font = .ReadProperty("Font", UserControl.Parent.Font)
        m_ActiveForeColor = .ReadProperty("ActiveForeColor", &H80000012)
        m_HoverColor = .ReadProperty("HoverColor", &H80000010)
        m_GradientEnd = .ReadProperty("GradientEnd", &HFFFFFF)
        m_GradientStart = .ReadProperty("GradientStart", &HFFC0C0)
        m_GradientStyle = .ReadProperty("GradientStyle", tsNoGradient)
        UserControl.shpHover.BorderColor = m_HoverColor
        m_InActiveForeColor = .ReadProperty("InActiveForeColor", &H80000011)
        m_Separators = .ReadProperty("Separators", True)
        m_SeparatorColor = .ReadProperty("SeparatorColor", &H80000010)
        m_TabBackColor = .ReadProperty("TabBackColor", &H8000000F)
        m_TabCount = .ReadProperty("TabCount", 3)
        m_TabHeight = .ReadProperty("TabHeight", 20)
        m_TabStyle = .ReadProperty("TabStyle", ts3D)
    End With
    Call HandleContainedControls(m_ActiveTab)
End Sub

Private Sub UserControl_Resize()
    
    With UserControl
        If (m_Width <> .ScaleWidth) Or (m_Height <> .ScaleHeight) Then
            m_Width = .ScaleWidth
            m_Height = .ScaleHeight
            Call MoveControls(0)
        End If
    End With
    
End Sub

Private Sub UserControl_Show()
    '   Reset the TabCount....when were in design this seems to reset,
    '   so we simply set it again which call the subs and redraws the
    '   control with the correct number of TabButtons
    TabCount = m_TabCount
    Call Refresh(m_ActiveTab)
End Sub

Private Sub UserControl_Terminate()
    '   Make sure the control is set to the initial tab or the Left values
    '   for the controls will be wrong....and the controls for Tab 0 will
    '   show up on the tab which was set when saved....
    ActiveTab = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("ActiveTab", m_ActiveTab, 0)
        Call .WriteProperty("Appearance", m_Appearance, tsFlat)
        m_Appearance = UserControl.Appearance
        Call .WriteProperty("BackColor", m_BackColor, &H8000000F)
        m_BackColor = UserControl.BackColor
        Call .WriteProperty("BorderStyle", m_BorderStyle, tsNone)
        m_BorderStyle = UserControl.BorderStyle
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", m_Font, UserControl.Parent.Font)
        Call .WriteProperty("ActiveForeColor", m_ActiveForeColor, &H80000012)
        Call .WriteProperty("HoverColor", m_HoverColor, &H80000010)
        Call .WriteProperty("GradientEnd", m_GradientEnd, &HFFFFFF)
        Call .WriteProperty("GradientStart", m_GradientStart, &HFFC0C0)
        Call .WriteProperty("GradientStyle", m_GradientStyle, tsNoGradient)
        m_HoverColor = UserControl.shpHover.BorderColor
        Call .WriteProperty("InActiveForeColor", m_InActiveForeColor, &H80000011)
        Call .WriteProperty("Separators", m_Separators, True)
        Call .WriteProperty("SeparatorColor", m_SeparatorColor, &H80000010)
        Call .WriteProperty("TabBackColor", m_TabBackColor, &H8000000F)
        Call .WriteProperty("TabCount", m_TabCount, 3)
        Call .WriteProperty("TabHeight", m_TabHeight, 20)
        Call .WriteProperty("TabStyle", m_TabStyle, ts3D)
    End With
End Sub

Public Property Get Version(Optional ByVal bDateTime As Boolean) As String
    ' The Version property
    If bDateTime Then
        Version = Major & "." & Minor & "." & Revision & " (" & DateTime & ")"
    Else
        Version = Major & "." & Minor & "." & Revision
    End If
End Property


