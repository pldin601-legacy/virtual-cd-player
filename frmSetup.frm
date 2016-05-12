VERSION 5.00
Begin VB.Form frmSetup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00004040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAutoLD 
      BackColor       =   &H00000000&
      Caption         =   "Автозапуск"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   915
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.CheckBox chkRand 
      BackColor       =   &H00000000&
      Caption         =   "Програвати випадковим чином"
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CheckBox chkSel 
      BackColor       =   &H00000000&
      Caption         =   "Програвати лише виділені записи"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Value           =   2  'Grayed
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Автоматичний старт для нових дисків"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   1740
      Width           =   2865
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shuffle"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3360
      TabIndex        =   5
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RENNSoft Multimedia (C) 2001"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4425
      TabIndex        =   3
      Top             =   2400
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Програвати лише ті записи, які позначені галочкою"
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   3360
      TabIndex        =   2
      Top             =   660
      Width           =   3195
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ОПЦІЇ ПРОГРАМИ"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   2145
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BB As Integer
Sub REPA()
On Error Resume Next
MX = Fix(Me.Width / Screen.TwipsPerPixelX)
For X = 0 To MX
  Me.Line (X * Screen.TwipsPerPixelX, 0)-(X * Screen.TwipsPerPixelX, Me.Height), RGB((150 / MX * (MX - X)), (150 / MX * (MX - X)), 0)
Next

Me.Line (0, 0)-(Me.Width, 0), RGB(255, 255, 0)
Me.Line (0, Me.Height - 15)-(Me.Width - 15, Me.Height - 15), RGB(155, 155, 0)
Me.Line (Me.Width - 15, 0)-(Me.Width - 15, Me.Height - 15), RGB(155, 155, 0)
Me.Line (0, 0)-(0, Me.Height - 15), RGB(255, 255, 0)


End Sub

Private Sub btnCancel_Click()

Unload Me

End Sub

Private Sub btnOK_Click()

On Error Resume Next

I = FreeFile

Open LowPath(App.Path) + App.EXEName + ".ini" For Output As #I

Print #I, "SETUP_PLAY_SELECTED "; Str(Me.chkSel.Value)
Print #I, "SETUP_PLAY_RANDOMIZED "; Str(Me.chkRand.Value)
Print #I, "SETUP_PLAY_LD "; Str(Me.chkAutoLD.Value)

Close #I

cdMain.LoadSetup

Unload Me

End Sub

Private Sub Form_Resize()
Dim MX, X
Dim SPS As Integer, SPR As Integer, SPL As Integer
On Error Resume Next

frmSetup.Top = cdMain.Top + cdMain.Height
frmSetup.Left = cdMain.Left + ((cdMain.Width - frmSetup.Width) / 2)

SPS = Val(GetIniRecord("SETUP_PLAY_SELECTED ", LowPath(App.Path) + App.EXEName + ".ini"))
SPR = Val(GetIniRecord("SETUP_PLAY_RANDOMIZED ", LowPath(App.Path) + App.EXEName + ".ini"))
SPL = Val(GetIniRecord("SETUP_PLAY_LD ", LowPath(App.Path) + App.EXEName + ".ini"))

chkSel.Value = SPS
chkRand.Value = SPR
chkAutoLD.Value = SPL

End Sub


