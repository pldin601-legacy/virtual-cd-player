VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl CDTrack 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ScaleHeight     =   2955
   ScaleWidth      =   3810
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   720
   End
   Begin PicClip.PictureClip PClip1 
      Left            =   840
      Top             =   1500
      _ExtentX        =   2037
      _ExtentY        =   582
      _Version        =   393216
      Cols            =   7
      Picture         =   "CDTrack.ctx":0000
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   180
      Picture         =   "CDTrack.ctx":1442
      ScaleHeight     =   330
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   0
      Width           =   165
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   0
      Picture         =   "CDTrack.ctx":179C
      ScaleHeight     =   330
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   0
      Width           =   165
   End
   Begin PicClip.PictureClip DigRes 
      Left            =   2460
      Top             =   1200
      _ExtentX        =   1455
      _ExtentY        =   1164
      _Version        =   393216
      Rows            =   2
      Cols            =   5
      Picture         =   "CDTrack.ctx":1AF6
   End
End
Attribute VB_Name = "CDTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Property Let TrackX(Track As Integer)
Dim X As String, M


X = Format$(Track, "00")
Dig1(0).Picture = DigRes.GraphicCell(Val(Mid$(X, 2, 1)))
If Val(Mid$(X, 1, 1)) = "0" Then
    Dig1(1).Picture = PClip1.GraphicCell(6)
Else
    Dig1(1).Picture = DigRes.GraphicCell(Val(Mid$(X, 1, 1)))
End If

End Property


