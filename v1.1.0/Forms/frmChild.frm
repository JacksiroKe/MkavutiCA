VERSION 5.00
Begin VB.Form frmChild 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Document"
   ClientHeight    =   3120
   ClientLeft      =   5940
   ClientTop       =   4410
   ClientWidth     =   4845
   ControlBox      =   0   'False
   Icon            =   "frmChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   4845
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTest 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Text            =   "Text"
      ToolTipText     =   "Changing Text will change the Form's Caption...."
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub txtTest_Change()
    Me.Caption = txtTest.Text
End Sub

