VERSION 5.00
Begin VB.Form frmEe_Control 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmEe_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long

Private Sub Form_Unload(Cancel As Integer)
    sndplaysound (App.Path & "\Tools\start.wav"), 1
End Sub

