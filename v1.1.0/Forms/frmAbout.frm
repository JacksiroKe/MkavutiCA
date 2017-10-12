VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Mkavuti Cyber Assistant"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   10815
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin MkavutiCyberAssistant.xpButton cmdClose 
      Height          =   735
      Left            =   8040
      TabIndex        =   1
      Top             =   5760
      Width           =   2055
      _ExtentX        =   3625
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
      MICON           =   "frmAbout.frx":0ECA
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ephantus Kiptanui - 0719 828 089"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   6120
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Concept by: Jackson Siro - 0711 474 787"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   5760
      Width           =   5655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "v 0.0.7"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   6825
      Left            =   0
      Picture         =   "frmAbout.frx":0EE6
      Top             =   0
      Width           =   10830
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1200
      X2              =   6000
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   240
      X2              =   5040
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00F0F0F0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2009 Juned Chhipa. All Rights Reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   8640
      Width           =   2520
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub
