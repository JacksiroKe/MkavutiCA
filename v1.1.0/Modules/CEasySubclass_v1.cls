VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEasySubclass_v1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'*    Author      : Paul Caton, Vlad Vissoultchev, modified by Andrea Batina[Revelatek]
'*    Date        : 30 July 2004
'*
'*    Component   : CEasySubclass_v1
'*    Description : Module-less, IDE safe, machine code subclassing thunk.
'*
'*    Dependencies: Subclassing/hooking sink interfaces 1.0 TypeLib [SubclassingSink.tlb]
'*
'*    Copyright   : Copyright � 2004 Revelatek. All rights reserved.
'*------------------------------------------------------------------------------
'*    Portions copyright (c) 2002 by Paul Caton <Paul_Caton@hotmail.com>
'*    Portions copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
'*******************************************************************************

Option Explicit

'////////////////////////////////////////////////////////////////////
'// Private/Public Type Definitions
Private Type OSVERSIONINFO
    dwOSVersionInfoSize             As Long
    dwMajorVersion                  As Long
    dwMinorVersion                  As Long
    dwBuildNumber                   As Long
    dwPlatformID                    As Long
    szCSDVersion                    As String * 128
End Type
Private Type UcsData
    hWnd                            As Long
    OrigWndProc                     As Long
    SinkInterface                   As Long
    MsgBuffer                       As Long
    BeforeBufferSize                As Long
    AfterBufferSize                 As Long
    AddrCallWindowProc              As Long
    AddrSetWindowLong               As Long
    AddrEbMode                      As Long
    AddrHeapFree                    As Long
    ProcessHeap                     As Long
End Type
Private Type UcsThunk
    Code(0 To &H190 \ 4 - 1)     As Long
    Data                            As UcsData
End Type

'////////////////////////////////////////////////////////////////////
'// Private/Public Win32 API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'////////////////////////////////////////////////////////////////////
'// Private/Public Constant Declarations
Private Const VER_PLATFORM_WIN32_NT     As Long = 2

Private Const STR_ASM_OPCODES           As String = "&H83EC8B55 &HE860F0C4 &H0 &HCEB815B &H33004010 &HF84589C0 &H8BFC4589 &H4011A08B &HFC98500 &H9684 &HFFF98300 &HBB8B1174 &H40119C &HF20C458B &H80850FAF &H83000000 &H4011B0BB &H36740000 &H11B093FF &HF8830040 &HC72B7502 &H1F845 &H63EB0000 &H1E75C085 &H1194B3FF &HFC6A0040 &H1190B3FF &H93FF0040 &H4011AC &H119883C7 &H40 &H938B0000 &H401198 &H3774D285 &HC085028B &H253174 &H75800000 &H458D532A &H458D5014 &H458D5010 &H458D500C &H458D5008 &H458D50FC &H8B5250F8 &H2050FF02 &HF87D835B &H8B850F00 &H53000000 &HFF1475FF &H75FF1075 " & _
                                                    "&H875FF0C &H1194B3FF &H93FF0040 &H4011A8 &HFC45895B &HF87D83 &H8B8B6875 &H4011A4 &H5E74C985 &H74FFF983 &H9CBB8B16 &H8B004011 &H4011A083 &H873C8D00 &HF20C458B &H834375AF &H4011B0BB &HB740000 &H11B093FF &HF8830040 &H8B2F7402 &H40119893 &H74D28500 &H85028B25 &H251F74C0 &H80000000 &HFF531875 &H75FF1475 &HC75FF10 &H8D0875FF &H5250FC45 &H50FF028B &H7D815B1C &H820C &HC7357500 &H40119083 &H0 &H838D00 &H50004010 &HB3FF006A &H4011B8 &H50F0458D &H1188838B &H45890040 &H8C838BF0 &H89004011 &HA3FFF445 &H4011B4 &HFC458B61 &H10C2C9"
Private Const STR_MODULE_USER32         As String = "user32"
Private Const STR_MODULE_KERNEL32       As String = "kernel32"
Private Const STR_MODULE_VBA6           As String = "vba6"
Private Const STR_MODULE_VBA5           As String = "vba5"
Private Const STR_CALLWINDOWPROC        As String = "CallWindowProcA"
Private Const STR_SETWINDOWLONG         As String = "SetWindowLongA"
Private Const STR_EBMODE                As String = "EbMode"
Private Const STR_HEAPFREE              As String = "HeapFree"

'////////////////////////////////////////////////////////////////////
'// Private/Public Variable Declarations
Private m_uThunk                    As UcsThunk
Private m_pThunk                    As Long
Private m_aBeforeMsgs()             As Long
Private m_aAfterMsgs()              As Long
Private m_vTag                      As Variant
Private m_oSinkInterface            As ISubclassingSink
Private m_bDontFree                 As Boolean

'//////////////////////////////////////////////////////////////////////////////
'//// PUBLIC PROPERTIES
'//////////////////////////////////////////////////////////////////////////////
Property Get ThunkAddress() As Long
    If m_pThunk = 0 Then m_pThunk = HeapAlloc(GetProcessHeap(), 0, Len(m_uThunk))
    ThunkAddress = m_pThunk
End Property
Property Get Tag() As Variant
    If IsObject(m_vTag) Then
        Set Tag = m_vTag
    Else
        Tag = m_vTag
    End If
End Property
Property Let Tag(vValue As Variant)
    m_vTag = vValue
End Property
Property Set Tag(ByVal oValue As Object)
    Set m_vTag = oValue
End Property
Property Get hWnd() As Long
    hWnd = m_uThunk.Data.hWnd
End Property

'//////////////////////////////////////////////////////////////////////////////
'//// PUBLIC METHODS
'//////////////////////////////////////////////////////////////////////////////
Public Function AddBeforeMsgs(ParamArray uMsgs()) As Boolean
    Dim lIdx            As Long
    
    AddBeforeMsgs = True
    For lIdx = 0 To UBound(uMsgs)
        AddBeforeMsgs = AddBeforeMsgs And pvAddMsg(m_aBeforeMsgs, uMsgs(lIdx))
    Next
End Function
Public Function AddAfterMsgs(ParamArray uMsgs()) As Boolean
    Dim lIdx            As Long
    
    AddAfterMsgs = True
    For lIdx = 0 To UBound(uMsgs)
        AddAfterMsgs = AddAfterMsgs And pvAddMsg(m_aAfterMsgs, uMsgs(lIdx))
    Next
End Function
Public Function Subclass(ByVal hWnd As Long, ByVal Sink As ISubclassingSink, Optional ByVal WeakReference As Boolean = False, Optional ByVal DontFree As Boolean = False) As Boolean
    With m_uThunk.Data
        '--- state check
        If .hWnd <> 0 Then
            Exit Function
        End If
        m_bDontFree = DontFree
        '--- store hWnd
        .hWnd = hWnd
        '--- store a reference (AddRef'd)
        If Not WeakReference Then
            Set m_oSinkInterface = Sink
        End If
        CopyMemory VarPtr(.SinkInterface), VarPtr(Sink), 4
        '--- store API functions entry points
        .AddrCallWindowProc = pvGetProcAddr(STR_MODULE_USER32, STR_CALLWINDOWPROC)
        .AddrSetWindowLong = pvGetProcAddr(STR_MODULE_USER32, STR_SETWINDOWLONG)
        '--- first try VBA6.DLL for EbMode function
        .AddrEbMode = pvGetProcAddr(STR_MODULE_VBA6, STR_EBMODE)
        '--- then VBA5.DLL
        If .AddrEbMode = 0 Then
            .AddrEbMode = pvGetProcAddr(STR_MODULE_VBA5, STR_EBMODE)
        End If
        '--- store heap management vars
        .AddrHeapFree = pvGetProcAddr(STR_MODULE_KERNEL32, STR_HEAPFREE)
        .ProcessHeap = GetProcessHeap()
        '--- change wndproc
        .OrigWndProc = SetWindowLong(hWnd, GWL_WNDPROC, ThunkAddress)
    End With
    '--- refresh heap chunk
    CopyMemory ThunkAddress, VarPtr(m_uThunk), Len(m_uThunk)
    '--- success
    Subclass = pvRefreshMsgsBuffer
End Function
Public Function UnSubclass() As Boolean
    Dim hSaveWnd            As Long
    
    With m_uThunk.Data
        ' Double subclass
        Debug.Assert GetWindowLong(.hWnd, GWL_WNDPROC) = 0 Or GetWindowLong(.hWnd, GWL_WNDPROC) = ThunkAddress
        
        '--- state check
        If .hWnd = 0 Then Exit Function
        '--- if stored reference is Release'd
        Set m_oSinkInterface = Nothing
        .SinkInterface = 0
        '--- prevent message buffers being traversed
        .BeforeBufferSize = 0
        .AfterBufferSize = 0
        '--- free previous buffer
        If .MsgBuffer <> 0 Then
            HeapFree GetProcessHeap(), 0, .MsgBuffer
            .MsgBuffer = 0
        End If
        '--- try to unsubclass
        If GetWindowLong(.hWnd, GWL_WNDPROC) = ThunkAddress Then
            SetWindowLong .hWnd, GWL_WNDPROC, .OrigWndProc
            If Not m_bDontFree Then
                HeapFree GetProcessHeap(), 0, m_pThunk
                m_pThunk = 0
            End If
        End If
        '--- can call Subclass later yet again
        hSaveWnd = .hWnd
        .hWnd = 0
    End With
    '--- if heap chunk available
    If IsWindow(hSaveWnd) And m_pThunk <> 0 Then
        If m_bDontFree And Not IsNT Then
            m_uThunk.Data.ProcessHeap = 0
        End If
        '--- inactivate heap chunk
        CopyMemory m_pThunk, VarPtr(m_uThunk), Len(m_uThunk)
        m_pThunk = 0
    End If
    '--- success
    UnSubclass = True
End Function
Public Function CallOrigWndProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    If m_uThunk.Data.hWnd <> 0 Then
        CallOrigWndProc = CallWindowProc(m_uThunk.Data.OrigWndProc, m_uThunk.Data.hWnd, uMsg, wParam, lParam)
    End If
End Function

'//////////////////////////////////////////////////////////////////////////////
'//// PRIVATE METHODS
'//////////////////////////////////////////////////////////////////////////////
Private Function pvAddMsg(aMsgs() As Long, ByVal uMsg As Long) As Boolean
    '--- if not filtered yet -> append msg
    If pvFindMsg(aMsgs, uMsg) < 0 Then
        '--- resize array
        If UBound(aMsgs) < 0 Then
            ReDim aMsgs(0 To 0)
        Else
            ReDim Preserve aMsgs(0 To UBound(aMsgs) + 1)
        End If
        '--- append new msg
        aMsgs(UBound(aMsgs)) = uMsg
        '--- success (or failure)
        pvAddMsg = pvRefreshMsgsBuffer()
    End If
End Function
Private Function pvFindMsg(aMsgs() As Long, ByVal uMsg As Long)
    Dim lIdx            As Long
    
    pvFindMsg = -1
    For lIdx = 0 To UBound(aMsgs)
        If aMsgs(lIdx) = uMsg Then
            pvFindMsg = lIdx
            Exit Function
        End If
    Next
End Function
Private Function pvRefreshMsgsBuffer() As Boolean
    Dim lBeforeSize     As Long
    Dim lAfterSize      As Long
    
    With m_uThunk.Data
        '--- init local vars
        lBeforeSize = UBound(m_aBeforeMsgs) + 1
        lAfterSize = UBound(m_aAfterMsgs) + 1
        '--- free previous buffer
        If .MsgBuffer <> 0 Then
            HeapFree GetProcessHeap(), 0, .MsgBuffer
            .MsgBuffer = 0
        End If
        '--- if any msg -> allocate new buffer
        If lBeforeSize + lAfterSize > 0 Then
            .MsgBuffer = HeapAlloc(GetProcessHeap(), 0, 4 * (lBeforeSize + lAfterSize))
            '--- fill new buffer: part 1
            If lBeforeSize > 0 Then
                CopyMemory .MsgBuffer, VarPtr(m_aBeforeMsgs(0)), 4 * lBeforeSize
            End If
            '--- fill new buffer: part 2
            If lAfterSize > 0 Then
                CopyMemory .MsgBuffer + 4 * lBeforeSize, VarPtr(m_aAfterMsgs(0)), 4 * lAfterSize
            End If
        End If
        .BeforeBufferSize = lBeforeSize
        .AfterBufferSize = lAfterSize
    End With
    '--- refresh heap chunk
    CopyMemory ThunkAddress, VarPtr(m_uThunk), Len(m_uThunk)
    '--- success
    pvRefreshMsgsBuffer = True
End Function
Private Function pvGetProcAddr(sModule As String, sFunction As String) As Long
    pvGetProcAddr = GetProcAddress(GetModuleHandle(sModule), sFunction)
End Function
Private Property Get IsNT() As Boolean
    Dim uVer            As OSVERSIONINFO
    uVer.dwOSVersionInfoSize = Len(uVer)
    If GetVersionEx(uVer) Then IsNT = uVer.dwPlatformID = VER_PLATFORM_WIN32_NT
End Property

'//////////////////////////////////////////////////////////////////////////////
'//// CLASS EVENTS
'//////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    Dim lIdx            As Long
    Dim vOpcode         As Variant
    
    '--- extract code
    For Each vOpcode In Split(STR_ASM_OPCODES)
        m_uThunk.Code(lIdx) = vOpcode
        lIdx = lIdx + 1
    Next
    '--- create "empty" arrays
    ReDim m_aBeforeMsgs(-1 To -1)
    ReDim m_aAfterMsgs(-1 To -1)
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    UnSubclass
End Sub