VERSION 5.00
Begin VB.Form frmBb_Computer 
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
   StartUpPosition =   2  'CenterScreen
   Begin MkavutiCyberAssistant.xpButton cmdUpdate 
      Height          =   615
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "frmBb_Computer.frx":0000
   End
   Begin MkavutiCyberAssistant.xpButton cmdDelete 
      Height          =   735
      Left            =   5520
      TabIndex        =   5
      Top             =   5400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
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
      MICON           =   "frmBb_Computer.frx":001C
   End
   Begin MkavutiCyberAssistant.xpButton cmdModify 
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Top             =   4440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
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
      MICON           =   "frmBb_Computer.frx":0038
   End
   Begin MkavutiCyberAssistant.xpButton cmdAddNow 
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "frmBb_Computer.frx":0054
   End
   Begin VB.ListBox lstComputer 
      Appearance      =   0  'Flat
      Height          =   5280
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "To Update or Delete a Computer Select it first on the list "
      Height          =   3135
      Left            =   5520
      TabIndex        =   6
      Top             =   960
      Width           =   2295
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
Attribute VB_Name = "frmBb_Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub cmdUpdate_Click()
    On Error GoTo ErrorHandler
        Set rs = New ADODB.Recordset
        rs.Open "Select * from my_clients where Computer='" & lstComputer.Text & "'", con, adOpenKeyset, adLockOptimistic
        rs.Update
        rs!Computer = txtName.Text
        rs.Update
        txtName.Text = ""
        Load_AllComputer
        frmDd_Customer.Computer_list
        cmdUpdate.Visible = False
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
        Exit Sub
ErrorHandler:
    'MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCc_Home.Enabled = True
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\Res\CyberCafe.mdb;"
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
        Set rs = New ADODB.Recordset
        rs.Open "Select * from my_clients", con, adOpenKeyset, adLockOptimistic
        rs.AddNew
        rs!Computer = txtName.Text
        rs.Update
        txtName.Text = ""
        Load_AllComputer
        frmDd_Customer.Computer_list
        Exit Sub
ErrorHandler:
    'MsgBox Err.Description & " No. " & Err.Number
    End If
    
End Sub

Private Sub Load_AllComputer()
    lstComputer.Clear
    Dim str As String
    On Error GoTo ErrorHandler
     Set rs = New ADODB.Recordset
        rs.Open "Select * from my_clients", con, adOpenKeyset, adLockOptimistic
        Do Until rs.EOF
            lstComputer.AddItem rs!Computer
            rs.MoveNext
        Loop
        rs.Close
        Exit Sub
ErrorHandler:
    'MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub cmdModify_Click()
    
    If lstComputer.ListCount > 0 Then
        cmdUpdate.Visible = True
        txtName.Text = lstComputer.Text
    End If
End Sub

Private Sub lstComputer_Click()
    If lstComputer.ListCount > 0 Then
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub txtName_Change()
    If txtName.Text = "" Then
        cmdAddNow.Enabled = False
    Else
        cmdAddNow.Enabled = True
    End If
End Sub

