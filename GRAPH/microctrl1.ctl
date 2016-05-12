VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl UserControl1 
   BackColor       =   &H00000040&
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   2895
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   180
      Picture         =   "microctrl1.ctx":0000
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox Micro 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   90
      Index           =   0
      Left            =   120
      Picture         =   "microctrl1.ctx":008A
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox Micro 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   90
      Index           =   1
      Left            =   60
      Picture         =   "microctrl1.ctx":0114
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox Micro 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   90
      Index           =   2
      Left            =   0
      Picture         =   "microctrl1.ctx":019E
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   0
      Width           =   60
   End
   Begin PicClip.PictureClip DigRes 
      Left            =   1860
      Top             =   2280
      _ExtentX        =   529
      _ExtentY        =   318
      _Version        =   327680
      Rows            =   2
      Cols            =   5
      Picture         =   "microctrl1.ctx":0228
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Let Value(Valn As Integer)

Dim X As String
X = Format$(Valn, "000")
Micro(0).Picture = DigRes.GraphicCell(Val(Mid$(X, 3, 1)))
Micro(1).Picture = DigRes.GraphicCell(Val(Mid$(X, 2, 1)))
Micro(2).Picture = DigRes.GraphicCell(Val(Mid$(X, 1, 1)))

End Property


