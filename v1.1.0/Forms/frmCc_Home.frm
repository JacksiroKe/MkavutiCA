VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00EBB85E-60A4-4EB4-A8A3-E451747B2506}#1.0#0"; "TABSMATA.OCX"
Begin VB.MDIForm frmCc_Home 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Mkavuti Cyber Assistant"
   ClientHeight    =   8235
   ClientLeft      =   2775
   ClientTop       =   2670
   ClientWidth     =   15555
   Icon            =   "frmCc_Home.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmCc_Home.frx":0ECA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin AsMdiTabs.TabSmata TabSmata 
      Left            =   6240
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   0
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   8400
      Top             =   720
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7740
      Left            =   0
      ScaleHeight     =   7740
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   0
      Width           =   4635
      Begin MkavutiCyberAssistant.xpButton cmdComputers 
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         TX              =   "&COMPUTERS"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCc_Home.frx":69E55
      End
      Begin MkavutiCyberAssistant.xpButton cmdCustomer 
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         TX              =   "&NEW CUSTOMER"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCc_Home.frx":69E71
      End
      Begin MkavutiCyberAssistant.xpButton cmdCustomers 
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         TX              =   "&CUSTOMERS"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   30
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCc_Home.frx":69E8D
      End
      Begin MkavutiCyberAssistant.xpButton cmdRecords 
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         TX              =   "&RECORDS"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   30
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCc_Home.frx":69EA9
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   3060
         Left            =   120
         TabIndex        =   5
         Top             =   5040
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5398
         _Version        =   393216
         ForeColor       =   0
         BackColor       =   0
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         MonthBackColor  =   8421504
         MultiSelect     =   -1  'True
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   84017153
         TitleBackColor  =   0
         TitleForeColor  =   16777215
         TrailingForeColor=   11183783
         CurrentDate     =   42695
         MaxDate         =   47848
         MinDate         =   42675
      End
      Begin MkavutiCyberAssistant.xpButton cmdPayment 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         TX              =   "&GET PAYMENT"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   30
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCc_Home.frx":69EC5
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7740
      Left            =   15360
      ScaleHeight     =   7740
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Width           =   200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   7740
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12718
            MinWidth        =   12348
            Text            =   "Mkavuti Cyber Assistant"
            TextSave        =   "Mkavuti Cyber Assistant"
            Object.ToolTipText     =   "Mkavuti Cyber Assistant"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            TextSave        =   "10/3/2017"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            TextSave        =   "9:04 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Customer"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPayment 
         Caption         =   "&Get Payment"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "&Customer View"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuRecords 
         Caption         =   "&Records View"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuComputer 
         Caption         =   "&Computer List"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuStyle 
      Caption         =   "&Tab Styles"
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style &1"
         Index           =   0
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style &2"
         Index           =   1
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style &3"
         Index           =   2
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style &4"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style &5"
         Index           =   4
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Tab Style &6"
         Index           =   5
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
Attribute VB_Name = "frmCc_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim ddate As Date, oldTime As Date, newTime As Date, diff As Date
Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long

Private Sub cmdComputers_Click()
    sndplaysound (App.Path & "\Res\Wavs\click.wav"), 1
    show_Computers_Window
End Sub

Private Sub cmdCustomer_Click()
    sndplaysound (App.Path & "\Res\Wavs\click.wav"), 1
    show_Customer_Window
End Sub

Private Sub cmdCustomers_Click()
    sndplaysound (App.Path & "\Res\Wavs\click.wav"), 1
    show_Customers_Window
End Sub

Private Sub cmdPayment_Click()
    Me.Enabled = False
    frmBb_Payment.Show , Me
End Sub

Private Sub cmdRecords_Click()
    sndplaysound (App.Path & "\Res\Wavs\click.wav"), 1
    show_Records_Window
End Sub

Private Sub TabSmata_ColorChanged(NewColor As stdole.OLE_COLOR)
    '  Picture1.BackColor = NewColor
    '  Where NewColor is a one of the color generated for OneNote style
End Sub

Private Sub TabSmata_DropdownButtonClick()
    PopupMenu mnuWindow
End Sub


Private Sub MDIForm_Load()
    oldTime = Time
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\CyberCafe.mdb;"
    con.Open
     
    mnuStyles_Click (3)
    show_Records_Window
    show_Customers_Window
   
    ddate = DateValue(Now)
    MonthView1.Value = ddate
    
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    'Shape1.Height = Me.Height - 1320 ';-)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    sndplaysound (App.Path & "\Res\Wavs\close.wav"), 1
    
End Sub

Private Sub mnuComputer_Click()
    show_Computers_Window
End Sub

Private Sub mnuCustomers_Click()
    show_Customers_Window
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileClose_Click()
    If ActiveForm Is Nothing Then Exit Sub
    Unload ActiveForm
End Sub

Private Sub mnuFileNew_Click()
    show_Customer_Window
End Sub

Private Sub mnuHelpAbout_Click()
    'TabSmata.About
End Sub

Private Sub mnuNavi_Click(Index As Integer)
Dim i As Long
    For i = 0 To 1
        mnuNavi(i).Checked = False
    Next i
    mnuNavi(Index).Checked = True
    TabSmata.NavigationStyle = Index
End Sub

Private Sub mnuPayment_Click()
    Me.Enabled = False
    frmBb_Payment.Show , Me
End Sub

Private Sub mnuPopupClose_Click()
    mnuFileClose_Click
End Sub

Private Sub show_Customer_Window()
    Dim frm As New frmDd_Customer
    frm.Show
End Sub

Private Sub show_Computers_Window()
    Dim frm As New frmBb_Computer
    frm.Show
    cmdComputers.Enabled = False
    mnuComputer.Enabled = False
End Sub

Private Sub show_Customers_Window()
    Dim frm As New frmDd_Customers
    frm.Show
    cmdCustomers.Enabled = False
    mnuCustomers.Enabled = False
End Sub

Private Sub show_Records_Window()
    Dim frm As New frmDd_Records
    frm.Show
    cmdRecords.Enabled = False
    mnuRecords.Enabled = False
End Sub

Private Sub TabSmata_TabBarClick(Button As Integer, x As Long, y As Long)
    Debug.Print "TabBarClick (" & Button & ", " & x & ", " & y & ")"
End Sub

Private Sub TabSmata_TabClick(TabHwnd As Long, Button As Integer, x As Long, y As Long)
    Debug.Print "TabClick (" & TabHwnd & ", " & Button & ", " & x & ", " & y & ")"
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuRecords_Click()
    show_Records_Window
End Sub

Private Sub mnuStyles_Click(Index As Integer)
Dim i As Long
    For i = 0 To 5
        mnuStyles(i).Checked = False
    Next i
    mnuStyles(Index).Checked = True
    TabSmata.Style = Index
End Sub

Function getSecsDiff(time_vl) As String
    Static x As Long
    Static zz$, ss$
    x = x + 1
    newTime = Time
    diff = DateDiff("s", time_vl, newTime)
    getSecsDiff = Format(((diff \ 60) * 0.5), "00.00")
End Function

Private Sub tmrUpdate_Timer()
    newTime = Time
    On Error GoTo ErrorHandler
        Set rs = New ADODB.Recordset
        rs.Open "Select * from my_clients where State=1", con, adOpenKeyset, adLockOptimistic
        Do Until rs.EOF
           rs!Time_Out = newTime
           rs!Cost_Kes = getSecsDiff(rs!Time_In)
           rs.Update
           rs.MoveNext
        Loop
        Exit Sub
ErrorHandler:
    'Computers_found = False
End Sub


