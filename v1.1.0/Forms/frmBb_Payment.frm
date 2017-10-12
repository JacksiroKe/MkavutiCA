VERSION 5.00
Begin VB.Form frmBb_Payment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Payment"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBb_Payment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstComputer 
      Appearance      =   0  'Flat
      Height          =   5280
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin MkavutiCyberAssistant.xpButton cmdPayment 
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   4920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmBb_Payment.frx":0ECA
   End
   Begin VB.Shape Shape7 
      Height          =   615
      Left            =   3960
      Top             =   480
      Width           =   5535
   End
   Begin VB.Shape Shape6 
      Height          =   615
      Left            =   3960
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Shape Shape5 
      Height          =   615
      Left            =   3960
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Left            =   3960
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   3960
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   3960
      Top             =   3960
      Width           =   5535
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount (Ksh): "
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3960
      TabIndex        =   13
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time: "
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out: "
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3960
      TabIndex        =   11
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time In: "
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3960
      TabIndex        =   10
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer: "
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3960
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date: "
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   480
      Width           =   2775
   End
   Begin VB.Line Line4 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   9480
      X2              =   3840
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lblAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Amount"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6720
      TabIndex        =   7
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblDuration 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Time"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6720
      TabIndex        =   6
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Line Line3 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   9480
      X2              =   3840
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   9480
      X2              =   3840
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblTimeout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Time Out"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6720
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblTimein 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Time In"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Customer"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6720
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   3840
      X2              =   3840
      Y1              =   480
      Y2              =   5760
   End
   Begin VB.Label lblDatein 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6720
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   5775
      Left            =   240
      Top             =   240
      Width           =   9495
   End
End
Attribute VB_Name = "frmBb_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long

Private Sub cmdPayment_Click()
On Error GoTo ErrorHandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from my_clients where Computer='" & lstComputer.Text & "'", con, adOpenKeyset, adLockOptimistic
    rs!Date_In = "0"
    rs!Customer = "my cust"
    rs!Time_In = "0"
    rs!Time_Out = "0"
    rs!Cost_Kes = "0"
    rs!State = "0"
    rs.Update
    AddToRecords
    
    lblDatein.Caption = ""
    lblCustomer.Caption = ""
    lblDatein.Caption = ""
    lblTimein.Caption = ""
    lblTimeout.Caption = ""
    lblAmount.Caption = ""
    
    lstComputer.Clear
    Load_Customers
    frmDd_Customers.Load_Customers
    sndplaysound (App.Path & "\Tools\paid.wav"), 1
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub AddToRecords()
On Error GoTo ErrorHandler
    Set rs = New ADODB.Recordset
    rs.Open "Select * from my_records", con, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!Computer = lstComputer.Text
    rs!Customer = lblCustomer.Caption
    rs!Date_In = lblDatein.Caption
    rs!Time_In = lblTimein.Caption
    rs!Time_Out = lblTimeout.Caption
    rs!Duration = lblDuration.Caption
    rs!Cost_Kes = lblAmount.Caption
    rs.Update
    Exit Sub
ErrorHandler:
    'MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sndplaysound (App.Path & "\Tools\close.wav"), 1
    frmCc_Home.Enabled = True
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\CyberCafe.mdb;"
    con.Open
    Load_Customers
End Sub
 
Private Sub Load_Customers()
    lstComputer.Clear
    Dim str As String
    On Error GoTo ErrorHandler
     Set rs = New ADODB.Recordset
        rs.Open "Select * from my_clients where State=1", con, adOpenKeyset, adLockOptimistic
        Do Until rs.EOF
            lstComputer.AddItem rs!Computer
            rs.MoveNext
        Loop
        rs.Close
        Exit Sub
ErrorHandler:
    'MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub lstComputer_Click()
cmdPayment.Enabled = True
On Error GoTo ErrorHandler
     Set rs = New ADODB.Recordset
        rs.Open "Select * from my_clients where Computer='" & lstComputer.Text & "'", con, adOpenKeyset, adLockOptimistic
        lblDatein.Caption = rs!Date_In
        lblCustomer.Caption = rs!Customer
        lblTimein.Caption = rs!Time_In
        lblTimeout.Caption = rs!Time_Out
        lblAmount.Caption = rs!Cost_Kes
        rs.Close
        Exit Sub
ErrorHandler:
    'MsgBox Err.Description & " No. " & Err.Number
End Sub
