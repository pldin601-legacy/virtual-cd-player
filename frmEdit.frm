VERSION 5.00
Begin VB.Form CDIE 
   BackColor       =   &H00004040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Реєстрація Компакт-диска"
   ClientHeight    =   3855
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1140
      TabIndex        =   6
      Text            =   "Untitled Disk"
      Top             =   3120
      Width           =   3315
   End
   Begin VB.TextBox txtTrack 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1140
      TabIndex        =   5
      Top             =   2760
      Width           =   3315
   End
   Begin VB.ListBox lstView 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbQSM 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RENNSOFT MULTIMEDIA"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2100
      TabIndex        =   9
      Top             =   3480
      Width           =   2355
   End
   Begin VB.Label txtID 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1140
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label LAB 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ІК Диска:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   3540
      Width           =   585
   End
   Begin VB.Label LAB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Назва запису:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   2820
      Width           =   900
   End
   Begin VB.Label LAB 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Назва диска:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   3180
      Width           =   855
   End
End
Attribute VB_Name = "CDIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next

lstView.Clear

Dim X, S, I
I = FreeFile
Me.Tag = cdMain.CDID.Tag
txtID = Format(Val(Me.Tag), "00000000")

For X = 1 To cdMain.MMCD.Tracks
  lstView.List(X - 1) = "Запис " + Format(X, "00")
Next X

X = 0

Open LowPath(App.Path) + Format(Val(Me.Tag), "00000000") + ".cd" For Input As #I

If Err = 0 Then
  Line Input #I, S
  txtTitle.Text = S
  Do
   X = X + 1
   Line Input #I, S
   lstView.List(X - 1) = S
  Loop While Not EOF(I)
End If

Close #I

lstView.ListIndex = 0


End Sub

Private Sub lstView_Click()
 txtTrack.Text = lstView.List(lstView.ListIndex)
 txtTrack.Tag = lstView.ListIndex
End Sub


Private Sub OKButton_Click()

On Error Resume Next

Dim X, S, I

I = FreeFile


Open LowPath(App.Path) + Format(Val(Me.Tag), "00000000") + ".cd" For Output As #I
Print #I, txtTitle.Text

For X = 1 To lstView.ListCount
  Print #I, lstView.List(X - 1)
Next X


Close #I

cdMain.UpdateCD

Unload Me

End Sub

Private Sub txtTitle_LostFocus()
 If txtTitle.Text = "" Then txtTitle.Text = "Untitled Disk"
End Sub


Private Sub txtTrack_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    txtTrack_LostFocus
  If lstView.ListIndex < lstView.ListCount - 1 Then
   lstView.ListIndex = lstView.ListIndex + 1
  Else
   lstView.ListIndex = 0
  End If
 txtTrack.SelStart = 0
 txtTrack.SelLength = Len(txtTrack.Text)
 KeyAscii = 0
 End If
End Sub


Private Sub txtTrack_LostFocus()
  If txtTrack.Text = "" Then txtTrack.Text = "Запис " + Format(Val(txtTrack.Tag), "00")
  lstView.List(Val(txtTrack.Tag)) = txtTrack.Text
End Sub

