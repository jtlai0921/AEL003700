VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "簡易計算器"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   4200
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command4 
      Caption         =   "除"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "乘"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "結束"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "減"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Text            =   " "
      Top             =   1200
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
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "結果"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "第二個數"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "第一個數"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   855
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
   Label4 = Val(Text1) - Val(Text2)
End Sub

Private Sub Command3_Click()
   Label4 = Val(Text1) * Val(Text2)
End Sub
Private Sub Command4_Click()
   Label4 = Val(Text1) / Val(Text2)
End Sub

Private Sub Command5_Click()
   End
End Sub

