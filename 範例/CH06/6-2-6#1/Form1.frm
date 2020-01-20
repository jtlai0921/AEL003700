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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   國文 = 80
   數學 = 85
   電腦 = 83
   平均 = (國文 + 數學 + 電腦) / 3
   Print "三科平均="; 平均
   Print "三科平均="; Format(平均, "###.##")  '取兩位小數,自動四捨五入
End Sub

