VERSION 5.00
Begin VB.UserControl XPProgressBar 
   BackColor       =   &H80000011&
   CanGetFocus     =   0   'False
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1650
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   1650
   ToolboxBitmap   =   "XpProgressBar.ctx":0000
   Begin VB.Image imgS 
      Height          =   240
      Index           =   5
      Left            =   0
      Picture         =   "XpProgressBar.ctx":0312
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image imgS 
      Height          =   240
      Index           =   4
      Left            =   0
      Picture         =   "XpProgressBar.ctx":03D4
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Image imgS 
      Height          =   240
      Index           =   3
      Left            =   0
      Picture         =   "XpProgressBar.ctx":0711
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image imgS 
      Height          =   240
      Index           =   2
      Left            =   0
      Picture         =   "XpProgressBar.ctx":0793
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image imgS 
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "XpProgressBar.ctx":0855
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image imgS 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "XpProgressBar.ctx":08D7
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image iP 
      Height          =   240
      Left            =   120
      Picture         =   "XpProgressBar.ctx":0959
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image iB 
      Height          =   240
      Left            =   960
      Picture         =   "XpProgressBar.ctx":09DB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30
   End
   Begin VB.Image iRight 
      Height          =   240
      Left            =   8760
      Picture         =   "XpProgressBar.ctx":0A9D
      Top             =   0
      Width           =   30
   End
   Begin VB.Image iLeft 
      Height          =   240
      Left            =   0
      Picture         =   "XpProgressBar.ctx":0B5F
      Top             =   0
      Width           =   30
   End
End
Attribute VB_Name = "XPProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum ProgressPicture
    XP_Green = 0: XP_Gold = 1: XP_Red = 2: XP_Blue = 3
End Enum
Dim mPic As ProgressPicture
Dim pValue As Single
Dim mAbout As String
Dim mEnabled As Boolean
Private Sub UserControl_Initialize()
    mAbout = "By Kundan"
    mEnabled = True
    iB.left = 30
    iP.left = 30
    UserControl.Height = 240
    UserControl.Width = 1260
    Me.Value = 0.7
    iRight.left = 1230
End Sub
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property
Public Property Let Enabled(ByVal what As Boolean)
    UserControl.Enabled() = what: mEnabled = what: checkEnabled
    PropertyChanged "Enabled"
End Property
Private Sub checkEnabled()
    If mEnabled = True Then
        mPic = mPic: MakeMeHappy (mPic)
        UserControl.BackColor = &H80000011
    Else
        iP.Picture = imgS(4).Picture
        iB.Picture = imgS(4).Picture
        UserControl.BackColor = &H8000000F
    End If
End Sub
Public Property Get Style() As ProgressPicture
    Style = mPic
End Property
Public Property Let Style(ByVal nStyle As ProgressPicture)
    If mEnabled Then mPic = nStyle: MakeMeHappy (mPic)
    PropertyChanged "Style"
End Property
Public Property Get About() As String
    About = mAbout
End Property
Public Property Let About(d As String)
    mAbout = d
    If mAbout <> "By Kundan" Then MsgBox "XP Progess Bar control" & vbNewLine & "By Kundan, IIT Delhi" & vbNewLine & vbNewLine & "Website: http://imkundan.tripod.com" & vbTab & vbNewLine & "Email: imkundan@yahoo.com" & vbNewLine & vbNewLine & "Please don't change it.", vbExclamation, "About.."
    mAbout = "By Kundan"
    PropertyChanged About
End Property
Private Sub MakeMeHappy(ch As Integer)
    iP.Picture = imgS(ch).Picture
    iB.Picture = imgS(5).Picture
End Sub
Public Property Get Value() As Single
    Value = pValue
End Property
Public Property Let Value(ByVal nValue As Single)
    On Error GoTo sex:
    If nValue > 1 Then GoTo er
    iP.Width = (UserControl.Width - 60) * nValue
    pValue = nValue
    PropertyChanged Value
    Exit Property
    
er:
    Err.Raise vbObjectError, , "overflow !"
sex:
Last.Timer2.Enabled = False
   
End Property
Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Height = 240
    iRight.left = UserControl.Width - 30
    iB.Width = UserControl.Width - 60
    iP.Width = (UserControl.Width) * pValue
    pValue = Value
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", 0.5)
    Style = PropBag.ReadProperty("Style", 0)
    About = PropBag.ReadProperty("About", "By Kundan")
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Style", mPic, XP_Green)
    Call PropBag.WriteProperty("Value", pValue, 0.5)
    Call PropBag.WriteProperty("About", mAbout, "By Kundan")
    Call PropBag.WriteProperty("Enabled", mEnabled, True)
End Sub
