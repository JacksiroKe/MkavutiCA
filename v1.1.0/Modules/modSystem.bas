Attribute VB_Name = "modSystem"
Option Explicit

'////////////////////////////////////////////////////////////////////
'// Private/Public Win32 API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'////////////////////////////////////////////////////////////////////
'// Private/Public Variable Declarations
Private m_oTimers       As New Collection   ' Timers collection
Private ExcludedChilds  As New Collection
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC

Public Property Get CaptionW(ByVal hWnd As Long) As String
    Dim lngLen As Long
    lngLen = DefWindowProcW(hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    If lngLen Then
        CaptionW = Space$(lngLen)
        DefWindowProcW hWnd, WM_GETTEXT, lngLen + 1, StrPtr(CaptionW)
    End If
End Property

Public Property Let CaptionW(ByVal hWnd As Long, ByVal NewValue As String)
    DefWindowProcW hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue)
End Property

'********************************************************************
'* Name: pEnumChildWindowProc
'* Description: Callback routine for enumerating MDI child windows.
'********************************************************************
Public Function pEnumChildWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim sBuf As String
    Dim sClass As String
    Dim iPos As Long
   
    If Not lParam = 0 Then
        sBuf = String$(261, 0)
        GetClassName hWnd, sBuf, 260
        iPos = InStr(sBuf, vbNullChar)
        If iPos > 1 Then
            sClass = Left$(sBuf, iPos - 1)
            If InStr(sClass, "Form") > 0 Then
                Dim ctlTab As TabSmata
                Dim oT As Object
                CopyMemory oT, lParam, 4
                Set ctlTab = oT
                CopyMemory oT, 0&, 4
                ctlTab.fAddMDIChildWindow hWnd
            End If
        End If
        pEnumChildWindowProc = 1
    End If
    
End Function

'********************************************************************
'* Name: TimerProc
'* Description: Timer callback method.
'********************************************************************
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    On Error Resume Next
    
    Dim oTimer As CTimer

    If hWnd = 0 Then
        ' Get timer object
        Set oTimer = m_oTimers.Item(CStr(idEvent))
        ' Raise timer event
        If Err.Number = 0 Then oTimer.RaiseTimerEvent
    End If
    
    Set oTimer = Nothing
End Sub

'********************************************************************
'* Name: AddTimer
'* Description: Add specified CTimer class into class collection.
'********************************************************************
Public Sub AddTimer(ByRef oTimer As CTimer, ByVal lTimerID As Long)
    On Error Resume Next
    
    m_oTimers.Add oTimer, CStr(lTimerID)
End Sub

'********************************************************************
'* Name: RemoveTimer
'* Description: Remove specified CTimer class from class collection.
'********************************************************************
Public Sub RemoveTimer(ByVal lTimerID As Long)
    On Error Resume Next
    
    m_oTimers.Remove CStr(lTimerID)
End Sub



