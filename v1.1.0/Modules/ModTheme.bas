Attribute VB_Name = "ModTheme"

Option Explicit

Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, _
                                                          ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
                                                                ByVal dwMaxNameChars As Long, _
                                                                ByVal pszColorBuff As Long, _
                                                                ByVal cchMaxColorChars As Long, _
                                                                ByVal pszSizeBuff As Long, _
                                                                ByVal cchMaxSizeChars As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

'//-- Current Theme Name
Public m_sCurrentSystemThemename As String

'Determines If The Current Window is Themed
Public Function AppThemed() As Boolean

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

'Returns The current Windows Theme Name
Public Sub GetThemeName(lngHwnd As Long)
    Dim hTheme As Long
    Dim sShellStyle As String
    Dim sThemeFile As String
    Dim lPtrThemeFile As Long
    Dim lPtrColorName As Long

    Dim iPos As Long
    On Error Resume Next
    hTheme = OpenThemeData(lngHwnd, StrPtr("ExplorerBar"))
    If Not hTheme = 0 Then
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lPtrThemeFile = VarPtr(bThemeFile(0))
        ReDim bColorName(0 To 260 * 2) As Byte
        lPtrColorName = VarPtr(bColorName(0))
        GetCurrentThemeName lPtrThemeFile, 260, lPtrColorName, 260, 0, 0
        sThemeFile = bThemeFile
        iPos = InStr(sThemeFile, vbNullChar)
        If iPos > 1 Then
            sThemeFile = Left$(sThemeFile, iPos - 1)
        End If
        m_sCurrentSystemThemename = bColorName
        iPos = InStr(m_sCurrentSystemThemename, vbNullChar)
        If iPos > 1 Then
            m_sCurrentSystemThemename = Left$(m_sCurrentSystemThemename, iPos - 1)
        End If
        sShellStyle = sThemeFile
        For iPos = Len(sThemeFile) To 1 Step -1
            If (Mid$(sThemeFile, iPos, 1) = "\") Then
                sShellStyle = Left$(sThemeFile, iPos)
                Exit For
            End If
        Next iPos
        sShellStyle = sShellStyle & "Shell\" & m_sCurrentSystemThemename & "\ShellStyle.dll"
        CloseThemeData hTheme
    Else
        m_sCurrentSystemThemename = "Classic"
    End If

End Sub

