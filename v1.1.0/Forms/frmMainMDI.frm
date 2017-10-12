VERSION 5.00
Object = "*\A..\jcMDITabs\jcMDITabs.vbp"
Begin VB.MDIForm frmMainMDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "jcMDITabs Control Demonstration"
   ClientHeight    =   6690
   ClientLeft      =   2775
   ClientTop       =   2670
   ClientWidth     =   10110
   LinkTopic       =   "MDIForm1"
   Begin jc_MDITabs.jcMDITabs jcMDITabs1 
      Left            =   7680
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   0
      NavigationStyle =   1
   End
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   10110
      TabIndex        =   11
      Top             =   6435
      Width           =   10110
   End
   Begin VB.PictureBox Picture4 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5940
      Left            =   9495
      ScaleHeight     =   5940
      ScaleWidth      =   615
      TabIndex        =   10
      Top             =   495
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5940
      Left            =   0
      ScaleHeight     =   5940
      ScaleWidth      =   2775
      TabIndex        =   1
      Top             =   495
      Width           =   2780
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   9795
         Left            =   0
         ScaleHeight     =   9795
         ScaleWidth      =   2775
         TabIndex        =   2
         Top             =   0
         Width           =   2775
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000015&
            BorderStyle     =   3  'Dot
            Height          =   7815
            Left            =   15
            Top             =   0
            Width           =   2745
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F1F1&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMainMDI.frx":0000
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1815
            Index           =   6
            Left            =   120
            TabIndex        =   9
            Top             =   4680
            Width           =   2535
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   2
            X1              =   120
            X2              =   2280
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F1F1&
            BackStyle       =   0  'Transparent
            Caption         =   "How To Use"
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
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   8
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F1F1&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMainMDI.frx":009D
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1815
            Index           =   4
            Left            =   120
            TabIndex        =   7
            Top             =   2880
            Width           =   2535
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F1F1&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMainMDI.frx":0150
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1335
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F1F1&
            BackStyle       =   0  'Transparent
            Caption         =   "Tabbed MDI user interfaces offer a number of improvements over traditional MDI interfaces."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   2535
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   120
            X2              =   2280
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   0
            X1              =   120
            X2              =   2280
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F1F1&
            BackStyle       =   0  'Transparent
            Caption         =   "More Information"
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
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   4320
            Width           =   2175
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F1F1&
            BackStyle       =   0  'Transparent
            Caption         =   "jcMDITabs"
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   2175
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10110
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   15240
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   15240
         X2              =   0
         Y1              =   10
         Y2              =   10
      End
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuStyle 
      Caption         =   "&Styles"
      Begin VB.Menu mnuStyles 
         Caption         =   "DotNET_1"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "DotNET_2"
         Index           =   1
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Office2000"
         Index           =   2
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Office2003"
         Index           =   3
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "OfficeOneNote"
         Index           =   4
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Whidbey (VS2005)"
         Index           =   5
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Properties"
      Begin VB.Menu mnuIcons 
         Caption         =   "Show Icons"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFocusrect 
         Caption         =   "Show FocusRect"
      End
   End
   Begin VB.Menu mnuNav 
      Caption         =   "&Navigation-Style"
      Begin VB.Menu mnuNavi 
         Caption         =   "Scroll Buttons"
         Index           =   0
      End
      Begin VB.Menu mnuNavi 
         Caption         =   "Dropdown Button"
         Checked         =   -1  'True
         Index           =   1
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub jcMDITabs1_ColorChanged(NewColor As stdole.OLE_COLOR)
    '--Here you can assign any control color --> NewColor
    '  for eg,
    '  Picture1.BackColor = NewColor
    '  Where NewColor is a one of the color generated for OneNote style
End Sub

Private Sub jcMDITabs1_DropdownButtonClick()
    PopupMenu mnuWindow
End Sub

Private Sub MDIForm_Load()
   'CaptionW(Me.hwnd) = "jcMDITabs Control Demonstration " & LoadResString(101)
   'CaptionW(frmChild.hwnd) = LoadResString(101)
   mnuFileNew_Click
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    Shape1.Height = Me.Height - 1320 ';-)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()

'On Error Resume Next
Static lDocID As Long
Dim frm As New frmChild

    lDocID = lDocID + 1
    frm.Caption = "Document " & lDocID
    'CaptionW(frm.hwnd) = LoadResString(101 + lDocID)
    frm.Show
    
End Sub

Private Sub mnuFileClose_Click()
    If ActiveForm Is Nothing Then Exit Sub
    Unload ActiveForm
End Sub

Private Sub mnuFocusrect_Click()
    mnuFocusrect.Checked = Not mnuFocusrect.Checked
    jcMDITabs1.ShowFocusRect = Not jcMDITabs1.ShowFocusRect
End Sub

Private Sub mnuHelpAbout_Click()
    jcMDITabs1.About
End Sub

Private Sub mnuIcons_Click()
    mnuIcons.Checked = Not mnuIcons.Checked
    jcMDITabs1.DrawIcons = Not jcMDITabs1.DrawIcons
End Sub

Private Sub mnuNavi_Click(Index As Integer)
Dim i As Long
    For i = 0 To 1
        mnuNavi(i).Checked = False
    Next i
    mnuNavi(Index).Checked = True
    jcMDITabs1.NavigationStyle = Index
End Sub

Private Sub mnuPopupClose_Click()
    mnuFileClose_Click
End Sub

Private Sub jcMDITabs1_TabBarClick(Button As Integer, x As Long, y As Long)
    Debug.Print "TabBarClick (" & Button & ", " & x & ", " & y & ")"
End Sub

Private Sub jcMDITabs1_TabClick(TabHwnd As Long, Button As Integer, x As Long, y As Long)
    Debug.Print "TabClick (" & TabHwnd & ", " & Button & ", " & x & ", " & y & ")"
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuStyles_Click(Index As Integer)
Dim i As Long
    For i = 0 To 5
        mnuStyles(i).Checked = False
    Next i
    mnuStyles(Index).Checked = True
    jcMDITabs1.Style = Index
End Sub

