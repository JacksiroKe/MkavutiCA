VERSION 5.00
Begin VB.Form frmEe_Payment 
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9150
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
   ScaleHeight     =   6345
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   645
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "frmEe_Payment"
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
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\CyberCafe.mdb;"
    con.Open
    Computer_list
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sndplaysound (App.Path & "\Tools\close.wav"), 1
    
End Sub


Public Sub Computer_list()

End Sub

