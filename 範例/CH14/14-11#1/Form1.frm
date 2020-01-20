VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "小精靈"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  '沒有框線
      Height          =   735
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '沒有框線
      Height          =   735
      Index           =   3
      Left            =   3120
      ScaleHeight     =   735
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '沒有框線
      Height          =   735
      Index           =   2
      Left            =   2160
      ScaleHeight     =   735
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '沒有框線
      Height          =   735
      Index           =   1
      Left            =   1200
      ScaleHeight     =   735
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '沒有框線
      Height          =   735
      Index           =   0
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Private Sub Form_Load()
   Timer1.Interval = 300
End Sub
Private Sub Form_Activate()
   For i = 0 To 3
      Picture1(i).Picture = LoadPicture("c:\mouse" + Mid(Str(i), 2, 1) + ".bmp")    '填滿
   Next i
End Sub
Private Sub Timer1_Timer()
   Picture2.Picture = Picture1(n).Picture
   Picture2.Left = Picture2.Left + Picture2.Width / 4
   If Picture2.Left > Form1.Width Then Picture2.Left = 0
   n = n + 1
   If n > 3 Then n = 0
End Sub
