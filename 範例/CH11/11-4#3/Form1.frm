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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Private Sub Form_Activate()
   Dim n(3) As String
   Dim s(3, 2) As Single, avg As Single
   For i = 1 To 3
      n(i) = InputBox$("請輸入第" & i & "個同學的姓名")
      s(i, 1) = Val(InputBox("請輸入第" & i & "個同學的國文成績"))
      s(i, 2) = Val(InputBox("請輸入第" & i & "個同學的數學成績"))
   Next i
   Print "姓名", "國文", "數學", "平均"
   Print
   For i = 1 To 3
      avg = (s(i, 1) + s(i, 2)) / 2
      Print n(i), s(i, 1), s(i, 2), avg
   Next i
End Sub

