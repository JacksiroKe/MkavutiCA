VERSION 5.00
Begin VB.Form frmDd_Customer 
   Caption         =   "New Customer"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8835
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDd_Customer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   8835
   WindowState     =   2  'Maximized
   Begin MkavutiCyberAssistant.xpButton cmdComputer 
      Height          =   615
      Left            =   4200
      TabIndex        =   8
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmDd_Customer.frx":000C
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   360
      Top             =   4320
   End
   Begin MkavutiCyberAssistant.xpButton cmdCustomer 
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   4440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1296
      TX              =   "Add Customer"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmDd_Customer.frx":0028
   End
   Begin VB.ComboBox cmbComputer 
      Height          =   645
      ItemData        =   "frmDd_Customer.frx":0044
      Left            =   4920
      List            =   "frmDd_Customer.frx":0046
      OLEDragMode     =   1  'Automatic
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   4320
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lblDateToday 
      Alignment       =   2  'Center
      Caption         =   "Date Today"
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   240
      Width           =   4575
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   8520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   8400
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblTimeNow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Time In:"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   8400
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   8400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Shape Shape1 
      Height          =   5415
      Left            =   120
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Name:"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Assign to Computer:"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   3855
   End
End
Attribute VB_Name = "frmDd_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim SavedThis As Boolean, Computers_found As Boolean
Dim ddate As Date, dtime As Date

Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\CyberCafe.mdb;"
    con.Open
    Computer_list
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sndplaysound (App.Path & "\Res\Wavs\close.wav"), 1
    
End Sub

Private Sub cmbComputer_Change()
    If cmbComputer.Text = "" Then
        cmdCustomer.Enabled = False
    Else
        cmdCustomer.Enabled = True
    End If
End Sub

Public Sub Computer_list()
On Error GoTo ErrorHandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from my_clients", con, adOpenKeyset, adLockOptimistic
    Do Until rs.EOF
        cmbComputer.AddItem rs!Computer
        rs.MoveNext
    Loop
    Exit Sub
ErrorHandler:
'Computers_found = False
End Sub

Private Sub cmbComputer_Click()
    cmbComputer_Change
End Sub

Private Sub cmdComputer_Click()
    frmCc_Home.Enabled = False
    frmBb_Computer.Show , frmCc_Home
End Sub

Private Sub cmdCustomer_Click()
    'On Error GoTo ErrorHandler
        sndplaysound (App.Path & "\Res\Wavs\welcome.wav"), 1
        Set rs = New ADODB.Recordset
        rs.Open "Select * from my_clients where Computer='" & cmbComputer.Text & "'", con, adOpenKeyset, adLockOptimistic
        rs!Customer = txtCustomer.Text
        rs!Time_In = lblTimeNow.Caption
        rs!Date_In = lblDateToday.Caption
        rs!client_state = "1"
        rs.Update
        rs.Close
        Unload Me
        Exit Sub
'ErrorHandler:
     'MsgBox "Unexpected error occurred, Please try again", vbExclamation, "Mkavuti Error"
End Sub

Private Function SiteSettings(option_title) As String
    Set rs = New ADODB.Recordset
    rs.Open "Select * from my_options WHERE title='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
    SiteSettings = rs!content
End Function

Private Function SaveSettings(option_title, option_cont) As Boolean
  On Error GoTo ErrorHandler
    Set rs = New ADODB.Recordset
        rs.Open "Select * from my_options where title ='" & option_title & "'", con, adOpenKeyset, adLockOptimistic
        rs!content = option_cont
        rs.Update
        rs.Close
        SaveSettings = True
        Exit Function
ErrorHandler:
 MsgBox "Unable to save changes. Either you obtain a fresh copy of vSongBook or contact the developer", vbExclamation, "vSongBook unexpected error"
End Function

Private Sub tmrTimer_Timer()
    ddate = DateValue(Now)
    dtime = TimeValue(Now)
    lblDateToday.Caption = ddate
    lblTimeNow.Caption = dtime
End Sub

