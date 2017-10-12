VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDd_Customers 
   Caption         =   "Cyber Customers"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10050
   ControlBox      =   0   'False
   Icon            =   "frmDd_Customers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   10020
      TabIndex        =   1
      Top             =   5055
      Width           =   10050
      Begin MkavutiCyberAssistant.xpButton cmdClose 
         Height          =   735
         Left            =   4440
         TabIndex        =   3
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1296
         TX              =   "&Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmDd_Customers.frx":000C
      End
      Begin MkavutiCyberAssistant.xpButton cmdPayment 
         Height          =   735
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1296
         TX              =   "Get &Payment"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmDd_Customers.frx":0028
      End
   End
   Begin VB.Timer tmrTimer 
      Interval        =   30000
      Left            =   120
      Top             =   2760
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   18441
      _Version        =   393216
      Rows            =   7
      Cols            =   10
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorSel    =   12632064
      ForeColorSel    =   255
      BackColorBkg    =   4210752
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      TextStyle       =   4
      TextStyleFixed  =   3
      GridLines       =   3
      SelectionMode   =   1
      AllowUserResizing=   3
      MousePointer    =   1
      FormatString    =   "Computer      |  Customer           |    Date_In         | Time_In    |  Cost_Kes "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDd_Customers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim SavedThis As Boolean
Dim ddate As Date, dtime As Date

Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPayment_Click()
    frmCc_Home.Enabled = False
    frmBb_Payment.Show , frmCc_Home
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\CyberCafe.mdb;"
    con.Open
     
    Load_Customers
End Sub

Public Sub Load_Customers()
On Error GoTo ErrorHandler
    
    Set rs = New ADODB.Recordset
    rs.Open "Select Computer, Customer, Date_In, Time_In, Cost_Kes from my_clients where State=1", con, adOpenStatic, adLockReadOnly
    
    MSFlexGrid1.Rows = rs.RecordCount + 1
    MSFlexGrid1.Cols = rs.Fields.Count
    MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
    MSFlexGrid1.Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    MSFlexGrid1.Row = 1
    Exit Sub
ErrorHandler:
    'MsgBox "Unexpected error occurred, Please try again", vbExclamation, "Mkavuti Error"
End Sub

Private Sub tmrTimer_Timer()
    Load_Customers
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    sndplaysound (App.Path & "\Res\Wavs\close.wav"), 1
    
    frmCc_Home.cmdCustomers.Enabled = True
    frmCc_Home.mnuCustomers.Enabled = True
End Sub
