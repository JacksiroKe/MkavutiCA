VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ComputerList"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8235
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin Mkavuti.xpButton cmdAddNow 
      Height          =   615
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      TX              =   "Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmBbComputer.frx":0000
   End
   Begin VB.ListBox lstComputer 
      Appearance      =   0  'Flat
      Height          =   5280
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   7695
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      Height          =   5535
      Left            =   120
      Top             =   840
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Computer Name:"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset


Private Sub Form_Unload(Cancel As Integer)
    frmCc_Home.Enabled = True
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Cyber_Cafe.mdb;"
    con.Open
    Load_AllComputer
End Sub
 
Private Sub cmdAddNow_Click()
    If txtName.Text = "" Then
        txtName.BackColor = &HFF&
        txtName.SetFocus
        Exit Sub
    Else
        On Error GoTo ErrorHandler
        Set Rs = New ADODB.Recordset
        Rs.Open "Select * from my_clients", con, adOpenKeyset, adLockOptimistic
        Rs.AddNew
        Rs!Computer = txtName.Text
        Rs.Update
        txtName.Text = ""
        Load_AllComputer
        frmDd_Customer.Computer_list
        Exit Sub
ErrorHandler:
    MsgBox Err.Description & " No. " & Err.Number
    End If
    
End Sub

Private Sub Load_AllComputer()
    lstComputer.Clear
    Dim str As String
    On Error GoTo ErrorHandler
     Set Rs = New ADODB.Recordset
        Rs.Open "Select * from my_clients", con, adOpenKeyset, adLockOptimistic
        Do Until Rs.EOF
            lstComputer.AddItem Rs!sb_title
            Rs.MoveNext
        Loop
        Rs.Close
        Exit Sub
ErrorHandler:
    MsgBox Err.Description & " No. " & Err.Number
End Sub


Private Sub txtName_Change()
    If txtName.Text = "" Then
        cmdAddNow.Enabled = False
    Else
        cmdAddNow.Enabled = True
    End If
End Sub

