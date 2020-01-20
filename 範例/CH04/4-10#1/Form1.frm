VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "加法計算器"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   4050
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "相加"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Text            =   " "
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   " "
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "結果"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "加數"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "被加數"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Label4 = Val(Text1) + Val(Text2)
End Sub

Private Sub Command2_Click()
   End
End Sub

