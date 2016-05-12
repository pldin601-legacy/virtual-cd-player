VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form cdMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Virtual Cool Disc Player"
   ClientHeight    =   2580
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   6765
   Icon            =   "frmPlay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCDen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   180
      ScaleHeight     =   195
      ScaleWidth      =   6375
      TabIndex        =   20
      Top             =   2280
      Width           =   6375
   End
   Begin VB.Timer tmrVar 
      Interval        =   250
      Left            =   5280
      Top             =   1380
   End
   Begin VB.CommandButton CDI 
      Caption         =   "CD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5220
      TabIndex        =   19
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton btnExit 
      Cancel          =   -1  'True
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      Top             =   1920
      Width           =   315
   End
   Begin VB.Timer tmrTrial 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5700
      Top             =   1380
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6120
      Top             =   1380
   End
   Begin VB.CommandButton btnMin 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5940
      TabIndex        =   15
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox Scroller 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   180
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   150
      TabIndex        =   13
      Top             =   1680
      Width           =   2955
   End
   Begin VB.CommandButton cmdUpLoad 
      Caption         =   "Перечитати"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   1920
      Width           =   1155
   End
   Begin VB.CommandButton cmdStore 
      Caption         =   "Конфіг"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3300
      TabIndex        =   11
      Top             =   1920
      Width           =   795
   End
   Begin VB.ListBox lstTracks 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1275
      IntegralHeight  =   0   'False
      Left            =   3300
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   540
      Width           =   3255
   End
   Begin MCI.MMControl MMCD 
      Height          =   300
      Left            =   180
      TabIndex        =   7
      Top             =   1860
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   529
      _Version        =   393216
      Frames          =   0
      BorderStyle     =   0
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      StopEnabled     =   -1  'True
      AutoEnable      =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "CDAudio"
      FileName        =   ""
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1035
      Left            =   180
      ScaleHeight     =   975
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   540
      Width           =   2955
      Begin VirtualCDP.TrackAll dspTrial 
         Height          =   135
         Left            =   2400
         TabIndex        =   17
         Top             =   660
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   238
      End
      Begin VirtualCDP.MTimer TrackLen 
         Height          =   330
         Left            =   1320
         Top             =   240
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
      End
      Begin VirtualCDP.CDTrack Track 
         Height          =   315
         Left            =   2460
         TabIndex        =   3
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
      End
      Begin VirtualCDP.TrackAll DiscLen 
         Height          =   135
         Left            =   1740
         TabIndex        =   2
         Top             =   660
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   238
      End
      Begin VirtualCDP.TrackAll DiscPos 
         Height          =   135
         Left            =   600
         TabIndex        =   1
         Top             =   660
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   238
      End
      Begin VirtualCDP.MTimer TrackPos 
         Height          =   330
         Left            =   180
         Top             =   240
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
      End
      Begin VB.Label CDID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disc ID: 0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   150
         Left            =   2340
         TabIndex        =   16
         Top             =   840
         Width           =   525
      End
      Begin VB.Label CAP 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ДИСК"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   120
         Index           =   4
         Left            =   1320
         TabIndex        =   9
         Top             =   645
         Width           =   300
      End
      Begin VB.Label CAP 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ДИСК"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   120
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   645
         Width           =   300
      End
      Begin VB.Label CAP 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ЗАПИС"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   120
         Index           =   2
         Left            =   2460
         TabIndex        =   6
         Top             =   60
         Width           =   360
      End
      Begin VB.Label CAP 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ТРИВАЛІСТЬ"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   120
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   60
         Width           =   675
      End
      Begin VB.Label CAP 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ПОЗИЦІЯ"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   120
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   60
         Width           =   465
      End
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RENNSoft Multimedia Software Kids"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   870
      TabIndex        =   14
      Top             =   60
      Width           =   5085
   End
End
Attribute VB_Name = "cdMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const SW_SHOW = 5
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40



Dim OldX, OldY
Dim Busy As Boolean
Dim ActiveID As String
Dim TTLTM

Dim SETUP_PLAY_SELECTED As Integer
Dim SETUP_PLAY_RANDOMIZED As Integer
Dim SETUP_PLAY_LD As Integer

Dim TrackStart As Integer
Dim TrackStop As Integer

Dim DSP_TRIAL As Integer

Dim ACT_TRACK As Integer
Sub LoadSetup()

On Error Resume Next

Dim SPS As Integer, SPR As Integer, SPL As Integer

SPS = Val(GetIniRecord("SETUP_PLAY_SELECTED ", LowPath(App.Path) + App.EXEName + ".ini"))
SPR = Val(GetIniRecord("SETUP_PLAY_RANDOMIZED ", LowPath(App.Path) + App.EXEName + ".ini"))
SPL = Val(GetIniRecord("SETUP_PLAY_LD ", LowPath(App.Path) + App.EXEName + ".ini"))

Let SETUP_PLAY_SELECTED = SPS
Let SETUP_PLAY_RANDOMIZED = SPR
Let SETUP_PLAY_LD = SPL

End Sub


Sub UpdateCD()
On Error Resume Next

lstTracks.Clear
Dim X, S
I = FreeFile

For X = 1 To MMCD.Tracks
  lstTracks.List(X - 1) = Format(X, "0") + ". Запис " + Format(X, "00")
Next X
X = 0

If MMCD.Tracks > 0 Then Title.Caption = "ДИСК НЕ ЗАРЕЄСТРОВАНИЙ"
If MMCD.Tracks = 0 Then Title.Caption = "VIRTUAL COOL DISC PLAYER. Ver " + GetVersion

Open LowPath(App.Path) + Format(Val(CDID.Tag), "00000000") + ".cd" For Input As #I

If Err = 0 Then
  Line Input #I, S
  Title.Caption = S
  Do
   X = X + 1
   Line Input #I, S
   lstTracks.List(X - 1) = Format(X, "0") + ". " + S
   lstTracks.Selected(X - 1) = True
  Loop While Not EOF(I)
End If

Close #I

ACT_TRACK = 0

ActiveID = CDID.Tag

End Sub


Sub UpdateSeek(TotalTime, MousePos As Single)

    MMCD.TimeFormat = 2

    Dim Seconds As Integer, Minutes As Integer
    Dim PercentSeek As Double, TSeconds As Integer

    PercentSeek = MousePos / Me.Scroller.ScaleWidth

    TSeconds = (Minute(TotalTime) + (Hour(TotalTime) * 60)) * PercentSeek
    Minutes = TSeconds \ 60
    Seconds = TSeconds Mod 60

    MMCD.Command = "Stop"
    MMCD.Command = "Seek"
    MMCD.Command = "Play"
End Sub


Private Sub btnExit_Click()
MMCD.Command = "Stop"
End
End Sub

Private Sub btnMin_Click()
Me.WindowState = 1
End Sub


Private Sub CDI_Click()

  If lstTracks.ListCount > 0 Then CDIE.Show

End Sub

Private Sub cmdStore_Click()

Dim DX, X, DY

If frmSetup.Visible = False Then

DX = frmSetup.Width
DY = frmSetup.Height

frmSetup.Height = 0
frmSetup.Width = 0
frmSetup.Show 0, Me

For X = 1 To 50
  frmSetup.Width = DX / 50 * X
  frmSetup.Height = DY / 50 * X
  DoEvents
Next

frmSetup.Width = DX
frmSetup.Height = DY
frmSetup.REPA
LoadSetup

End If


End Sub

Private Sub cmdUpload_Click()
UpdateCD

End Sub



Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Me.ScaleMode = vbPixels

MMCD.Wait = True
MMCD.Shareable = True
MMCD.Command = "Open"

LoadSetup

DSP_TRIAL = 80 * 60


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldX = X
OldY = Y

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ML, MT

If Button = 1 Then

 ML = Me.Left + (X - OldX)
 MT = Me.Top + (Y - OldY)
 
 If ML <= 120 Then ML = 0
 If MT <= 120 Then MT = 0
 
 If ML > Screen.Width - Me.Width - 120 Then ML = Screen.Width - Me.Width
 If MT > Screen.Height - Me.Height - 120 Then MT = Screen.Height - Me.Height
 
 Me.Left = ML
 Me.Top = MT

 X = ML
 Y = MT
 
  If frmSetup.Visible = True Then
    frmSetup.Top = cdMain.Top + cdMain.Height
    frmSetup.Left = Me.Left + ((Me.Width - frmSetup.Width) / 2)
  End If
 
End If

End Sub


Private Sub Form_Resize()

Me.Line (0, 0)-(Me.Width, 0), RGB(255, 255, 0)
Me.Line (0, Me.Height - 15)-(Me.Width - 15, Me.Height - 15), RGB(155, 155, 0)
Me.Line (Me.Width - 15, 0)-(Me.Width - 15, Me.Height - 15), RGB(155, 155, 0)
Me.Line (0, 0)-(0, Me.Height - 15), RGB(255, 255, 0)

Ret = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 20, 20)
SetWindowRgn Me.hwnd, Ret, True


End Sub

Private Sub lstTracks_DblClick()
 
  MMCD.TimeFormat = 10
  MMCD.To = lstTracks.ListIndex + 1
  MMCD.Command = "Stop"
  MMCD.Command = "Seek"
  MMCD.Command = "Play"

End Sub

Sub ScrChange(Max As Integer, Min As Integer)
Dim MOCbKA As Integer
Dim Valve As Integer
Dim Vise As Integer

Vise = Min

Scroller.Cls

For MOCbKA = Vise To Vise + 10
 Valve = MOCbKA - Vise
 Scroller.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(255 - (255 / 10 * Valve), 200 - (200 / 10 * Valve), 0), BF
Next

For MOCbKA = Vise - 10 To Vise
 Valve = MOCbKA - (Vise - 10)
 Scroller.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(100 / 10 * Valve, 255 / 10 * Valve, 0), BF
Next

Scroller.Line (Vise, 0)-(Vise + 1, 1), RGB(255, 255, 0), BF

End Sub

Private Sub MMCD_StatusUpdate()
DoEvents

End Sub

Private Sub MMCD_StopClick(Cancel As Integer)

MMCD.Command = "STOP"
MMCD.To = 1
MMCD.Command = "SEEK"

End Sub


Private Sub Scroller_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Interval = 0
Scroller_MouseMove Button, Shift, X, Y
End Sub


Private Sub Scroller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X > Scroller.ScaleWidth Then X = Scroller.ScaleWidth
If X < 0 Then X = 0

If xButton = 1 Then
  UpdateSeek TTLTM, X + TrackStart
End If

If Button = 1 Then
 MMCD.TimeFormat = 0
 Melisa = X * 1000
 aMins = Fix((Melisa / 1000) / 60)
 aSecs = Melisa / 1000 Mod 60
 DiscPos.TimeCDSet = Format$(aMins, "00") + ":" + Format$(aSecs, "00")
 ScrChange Scroller.ScaleWidth, CInt(X)
End If

End Sub


Private Sub Scroller_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
 If X > Scroller.ScaleWidth Then X = Scroller.ScaleWidth
 If X < 0 Then X = 0
 
 Dim aMins, aSecs
 Dim Melisa As Single
 Dim Track As Integer
 
 MMCD.TimeFormat = 0
 Melisa = (X + TrackStart) * 1000

 ScrChange Scroller.ScaleWidth, CInt(X)
   
 aMins = Fix((Melisa / 1000) / 60)
 aSecs = Melisa / 1000 Mod 60
 DiscPos.TimeCDSet = Format$(aMins, "00") + ":" + Format$(aSecs, "00")
 MMCD.To = (X + TrackStart) * 1000
 MMCD.Command = "Stop"
 MMCD.Command = "Seek"
 MMCD.Command = "Play"
 
End If

Timer2.Interval = 1000
End Sub


Private Sub Timer1_Timer()
End Sub

Private Sub Timer2_Timer()
DoEvents
On Error Resume Next

If ActiveID <> CDID.Tag Then
  MMCD.Command = "Close"
  MMCD.Command = "Open"
  UpdateCD
  If SETUP_PLAY_LD Then MMCD.Command = "Play"
End If

Busy = True

Dim CTrack


MMCD.TimeFormat = 10
MMCD.Track = MMCD.Position And &HFF
Track.TrackX = MMCD.Track


Dim lSecs, lMins, lSecls, pSecs, pMins, pSecls
MMCD.TimeFormat = 0
lSecls = Fix(MMCD.Length / 1000)
pSecls = Fix(MMCD.Position / 1000)
lMins = Fix(lSecls / 60)
pMins = Fix(pSecls / 60)
lSecs = lSecls Mod 60
pSecs = pSecls Mod 60

Dim TlSecs, TlMins, TlSecls, TpSecs, TpMins, TpSecls, TDSecls

MMCD.TimeFormat = 0

TlSecls = Fix(MMCD.TrackLength / 1000)
TDSecls = Fix(MMCD.Length / 1000)
MMCD.TimeFormat = 10

TpMins = (MMCD.Position And &HFF00&) \ &H100
TpSecs = (MMCD.Position And &HFF0000) \ &H10000

TTLTM = TimeSerial(0, TpMins, TpSecs)


TlMins = Fix(TlSecls / 60)
TlSecs = TlSecls Mod 60

TrackStart = ((pMins * 60) + pSecs) - ((TpMins * 60) + TpSecs)
TrackStop = TrackStart + TlSecls



Scroller.ScaleWidth = TlSecls

' TrackStart = 0
' TrackStop = lMins * 60 + lSecs
' Scroller.ToolTipText = "From " + Format(TimeSerial(0, 0, TrackStart), "M:ss") + " to " + Format(TimeSerial(0, 0, TrackStop), "M:ss")

TrackPos.TimeSet = Format(TpMins, "00") + ":" + Format(TpSecs, "00")
TrackLen.TimeSet = Format(TlMins, "00") + ":" + Format(TlSecs, "00")
DiscPos.TimeCDSet = Format(pMins, "00") + ":" + Format(pSecs, "00")
DiscLen.TimeCDSet = Format(lMins, "00") + ":" + Format(lSecs, "00")
 
If ACT_TRACK <> MMCD.Track Then
 ACT_TRACK = MMCD.Track
 If lstTracks.ListCount >= MMCD.Track Then
   lstTracks.ListIndex = MMCD.Track - 1
 End If
End If

MMCD.TimeFormat = 0
' ScrChange CInt(MMCD.Length / 1000), (TpMins * 60) + TpSecs
ScrChange CInt(MMCD.Length / 1000), CInt(MMCD.Position / 1000) - TrackStart

picCDen.ScaleWidth = TDSecls
picCDen.ScaleHeight = 1
picCDen.Cls

Dim N, U
U = 0
For N = 1 To MMCD.Tracks
  MMCD.Track = N
  U = U + CInt(MMCD.TrackLength / 1000)
  picCDen.Line (U, 0)-(U + 1, 1), RGB(0, 0, 255), BF
Next N

picCDen.Line (CInt(MMCD.Position / 1000), 0)-(CInt(MMCD.Position / 1000) + 1, 1), RGB(255, 255, 255), BF
 
If MMCD.Tracks > 1 Then CDID.Caption = "Disk ID: " + Format(MMCD.Length, "00000000")
If MMCD.Tracks <= 1 Then CDID.Caption = "Please insert an audio CD"

CDID.Tag = MMCD.Length

Me.Caption = lstTracks.List(ACT_TRACK - 1) + " [" + Format(TpMins, "0") + ":" + Format(TpSecs, "00") + "]"

Busy = False


End Sub

Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldX = X
OldY = Y

End Sub


Private Sub title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ML, MT

If Button = 1 Then

 ML = Me.Left + (X - OldX)
 MT = Me.Top + (Y - OldY)
 
 If ML <= 120 Then ML = 0
 If MT <= 120 Then MT = 0
 
 If ML > Screen.Width - Me.Width - 120 Then ML = Screen.Width - Me.Width
 If MT > Screen.Height - Me.Height - 120 Then MT = Screen.Height - Me.Height
 
 Me.Left = ML
 Me.Top = MT

 X = ML
 Y = MT
 
  If frmSetup.Visible = True Then
    frmSetup.Top = cdMain.Top + cdMain.Height
    frmSetup.Left = Me.Left + ((Me.Width - frmSetup.Width) / 2)
  End If
 
End If


End Sub


Sub PlayNextSel()
 Dim A, B
 
 For A = lstTracks.ListIndex + 1 To lstTracks.ListCount - 1
   If lstTracks.Selected(A) = True Then
      lstTracks.ListIndex = A
      lstTracks_DblClick
      Exit For
   End If
 Next
   
End Sub

Private Sub tmrTrial_Timer()

On Error Resume Next

Dim DP_SEC, DP_MIN

DSP_TRIAL = DSP_TRIAL - 1

DP_SEC = DSP_TRIAL Mod 60
DP_MIN = Fix(DSP_TRIAL / 60)

dspTrial.TimeCDSet = Format(DP_MIN, "00") + ":" + Format(DP_SEC, "00")

If DSP_TRIAL = -1 Then
  MsgBox "Усе! Час тріального користування даної програми закінчився. Якщо ви бажаєте користуватися програмою і на далі, то вам необхідно придбати ліцензійну версію програми, яка коштує приблизно 8 грн.", vbInformation, "Trial Version"
  MsgBox "Дякую за користування програмами софтмейк студії RENNSoft Multimedia Software!", vbOKOnly
  End
End If

End Sub


Private Sub tmrVar_Timer()
On Error Resume Next
If MMCD.Mode = 525 Then Exit Sub

MMCD.TimeFormat = 10
MMCD.Track = MMCD.Position And &HFF

If SETUP_PLAY_SELECTED = 1 Then
  If lstTracks.Selected(MMCD.Track - 1) = False Then PlayNextSel
End If


End Sub


