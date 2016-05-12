VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl TrackAll 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   660
   ScaleWidth      =   4800
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   0
      Left            =   360
      Picture         =   "TrackAll.ctx":0000
      ScaleHeight     =   120
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   1
      Left            =   270
      Picture         =   "TrackAll.ctx":00C2
      ScaleHeight     =   120
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   180
      Picture         =   "TrackAll.ctx":0184
      ScaleHeight     =   120
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   2
      Left            =   90
      Picture         =   "TrackAll.ctx":0246
      ScaleHeight     =   120
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   3
      Left            =   0
      Picture         =   "TrackAll.ctx":0308
      ScaleHeight     =   120
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   0
      Width           =   75
   End
   Begin PicClip.PictureClip DigResource 
      Left            =   960
      Top             =   120
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      Rows            =   2
      Cols            =   5
      Picture         =   "TrackAll.ctx":03CA
   End
End
Attribute VB_Name = "TrackAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Let TimeCDSet(TimeNow As String)
Dim Secs As String
Dim MinS As String

Secs = Mid$(TimeNow, 4, 2)
MinS = Mid$(TimeNow, 1, 2)

Dig1(0).Picture = DigResource.GraphicCell(Val(Mid$(Secs, 2, 1)))
Dig1(1).Picture = DigResource.GraphicCell(Val(Mid$(Secs, 1, 1)))
Dig1(2).Picture = DigResource.GraphicCell(Val(Mid$(MinS, 2, 1)))
Dig1(3).Picture = DigResource.GraphicCell(Val(Mid$(MinS, 1, 1)))

End Property

