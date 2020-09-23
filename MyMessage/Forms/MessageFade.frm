VERSION 5.00
Begin VB.Form MessageFade 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer MessageFade 
      Interval        =   50
      Left            =   0
      Top             =   1320
   End
   Begin VB.Timer FadeDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   1320
   End
   Begin VB.OptionButton FadeIn 
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton FadeOut 
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lDescription 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "MessageFade.frx":0000
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lClose 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   30
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "MessageFade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MESSAGE ON TOP
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

'GET SCREEN RIGHT CORNER
Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long

Private Const SPI_GETWORKAREA = 48

'FORM FADE
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Dim Fade As Integer

'FORM FADE
Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long

Dim Msg As Long
On Error Resume Next

If Perc < 0 Or Perc > 255 Then
  
    MakeTransparent = 1

Else
  
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
    MakeTransparent = 0

End If

If Err Then
  
      MakeTransparent = 2

End If

End Function

'MESSAGE ON TOP
Public Sub MessageOnTop(hWindow As Long, bTopMost As Boolean)
    
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    
Dim wFlags
Dim placement
    
wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
placement = HWND_TOPMOST
    
SetWindowPos hWindow, placement, 0, 0, 0, 0, wFlags

End Sub

'POSITION THE FORM IN THE RIGHT CORNER OF SCREEN
Private Sub PlaceMessageInLowerRight(ByVal frm As Form, ByVal right_margin As Single, ByVal bottom_margin As Single)

Dim wa_info As RECT

If MainForm.Taskbar.Value = True Then

    If SystemParametersInfo(SPI_GETWORKAREA, 0, wa_info, 0) <> 0 Then

        'GOT POSITION, PLACE THE FORM NOW
        frm.Left = ScaleX(wa_info.Right, vbPixels, vbTwips) - Width - right_margin
        frm.top = ScaleY(wa_info.Bottom, vbPixels, vbTwips) - Height - bottom_margin
    
    End If

End If

If MainForm.WTaskbar.Value = True Then
        
    'DID NOT GOT THE WORK AREA BOUNDS, USE THE ENTIRE SCREEN
    frm.Left = Screen.Width - Width - right_margin
    frm.top = Screen.Height - Height - bottom_margin
    
End If

End Sub

Private Sub FadeDelay_Timer()

FadeDelay.Interval = FadeDelay.Interval + 1000

If FadeDelay.Interval = 3000 Then

    FadeIn.Value = False
    FadeOut.Value = True
    FadeDelay.Enabled = False
    MessageFade.Enabled = True

End If

End Sub

Private Sub Form_Load()

Call MessageOnTop(Me.hwnd, True) 'MESSAGE ON TOP

PlaceMessageInLowerRight Me, 0, 0 'MESSAGE PLACEMENT

Fade = 0
MakeTransparent Me.hwnd, 0

End Sub

Private Sub lClose_Click()

FadeDelay.Enabled = False   'DONT ALLOW TO PROCEED TO DELAY TIME
FadeIn.Value = False        'SKIP MESSAGE FADE IN
FadeOut.Value = True        'ALLOW MESSAGE FADE OUT
MessageFade.Enabled = True  'KEEP MESSAGE FADING

End Sub

Private Sub lClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

lClose.ForeColor = vbWhite

End Sub

Private Sub lClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

lClose.ForeColor = vbBlack

End Sub

'FORM FADE IN EFFECT CONTROL TIMER
Private Sub messagefade_Timer()

'THE FADING EFFECT SPEED CAN BE CONTROLLED BY CHANGING THE MESSGEFADE INTERVAL AND / OR
'BY CHANGING THE BELOW ADDITION AND SUBRATION VALUE (i.e. 20)

If FadeIn.Value = True Then

    If Fade <= 255 Then
 
        MakeTransparent Me.hwnd, Fade
        Fade = Fade + 20                'FADE IN MESSAGE UNTIL IT IS FULLY APPEARED

    Else

        MessageFade.Enabled = False     'STOP WHEN FULLY APPEARED
        FadeDelay.Enabled = True        'START DELAY TIMER
        MakeTransparent Me.hwnd, 255    '255 = FULLY APPEARED

    End If

End If

If FadeOut.Value = True Then

    If Fade >= 0 Then
 
        MakeTransparent Me.hwnd, Fade
        Fade = Fade - 20                'FADE OUT MESSAGE UNTIL IT IS FULLY DISAPPEARED

    Else

        MessageFade.Enabled = False     'STOP WHEN FULLY DISAPPEARED
        MakeTransparent Me.hwnd, 0      '0 = FULLY DISAPPEARED
        On Error Resume Next
        Me.Hide                         'HIDE WHEN DONE

    End If

End If

End Sub


