VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   3
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   2280
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   2
      Left            =   1920
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   2280
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   1
      Left            =   1080
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   2280
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   0
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   Const pi = 3.141593
   For i = 0 To 3
      Picture1(i).FillColor = QBColor(0)   '黑色
      Picture1(i).FillStyle = 0            '填滿
      Picture1(i).Scale (0, 0)-(10, 10)  '設定內部座標
   Next i
   '分別在四個圖片方塊中畫小精靈
   Picture1(0).Circle (5, 5), 4, , -pi / 4, -pi * 7 / 4
   Picture1(1).Circle (5, 5), 4, , -pi / 6, -pi * 11 / 6
   Picture1(2).Circle (5, 5), 4
   Picture1(3).Circle (5, 5), 4, , -pi / 6, -pi * 11 / 6
   For i = 0 To 3   '將四個圖片方塊中的小精靈分別存檔
      SavePicture Picture1(i).Image, "c:\mouse" + Mid(Str(i), 2, 1) + ".bmp"
   Next i
End Sub

Private Sub Form_Load()
   For i = 0 To 3   '設定四個圖片方塊中的AutoRedraw屬性為True
      Picture1(i).AutoRedraw = True
   Next i
End Sub
