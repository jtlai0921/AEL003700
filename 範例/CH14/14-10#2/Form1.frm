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
   Rem 將四個小精靈檔案下載到四個圖片方塊中
   For i = 0 To 3
      Picture1(i).Picture = LoadPicture("c:\mouse" + _
                            Mid(Str(i), 2, 1) + ".bmp")
   Next i
End Sub

