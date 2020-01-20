VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯示"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '暫止
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入密碼："
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
      Print "你輸入的密碼是："; Text1
End Sub

Private Sub Command2_Click()
   End
End Sub
