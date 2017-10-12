VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAa_Splash 
   BorderStyle     =   0  'None
   Caption         =   "frmMain"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   Icon            =   "frmAa_Splash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAa_Splash.frx":0ECA
   ScaleHeight     =   6750
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pgbLoading 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   5760
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© 2016 Jackson Siro | Ephantus Kiptanui [KTTC, Gigiri, Nairobi, Kenya]"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   6120
      Width           =   7815
   End
End
Attribute VB_Name = "frmAa_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim shell As WshShell, lngReturnCode As Long
Dim strShellCmd As String, strShellCommand As String

Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long

Private Sub Form_Unload(Cancel As Integer)
    'sndplaysound (App.Path & "\Tools\door.wav"), 1

End Sub

' this will make ur progress bar Run
Private Sub Timer1_Timer()
    On Error GoTo ErrorHandler:
        With pgbLoading
            .Value = .Value + 1
        End With
Exit Sub
ErrorHandler:
    If Err.Number = 380 Then
        'sndplaysound (App.Path & "\Tools\door.wav"), 1

        Unload Me
        frmCc_Home.Show
    End If
End Sub


