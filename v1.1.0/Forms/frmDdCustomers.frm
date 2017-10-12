VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDdCustomers 
   Caption         =   "Cyber Customers"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
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
      Cols            =   5
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
      AllowUserResizing=   3
      MousePointer    =   1
      FormatString    =   "Computer      |  Customer           |    Date_In         | Time_In    |  Cost_Kes"
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
Attribute VB_Name = "frmDdCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Dim SavedThis As Boolean
Dim ddate As Date, dtime As Date

Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\CyberCafe.mdb;"
    con.Open
    brwCustomers.Navigate "about:blank"
    
    Load_Customers
End Sub

Private Sub Load_Customers()
On Error GoTo ErrorHandler
    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from my_clients where client_state='1'", con, adOpenStatic, adLockReadOnly
        
    
    Exit Sub
ErrorHandler:
    'MsgBox "Unexpected error occurred, Please try again", vbExclamation, "Mkavuti Error"
End Sub

Private Sub Form_Resize()
    brwCustomers.Width = frmDd_Customers.Width
    brwCustomers.Height = frmDd_Customers.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sndplaysound (App.Path & "\Tools\close.wav"), 1
    
    frmCc_Home.cmdCustomers.Enabled = True
    frmCc_Home.mnuCustomers.Enabled = True
End Sub

