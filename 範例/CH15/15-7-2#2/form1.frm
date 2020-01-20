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
Private Sub Form_Activate()
   Open "a:\score.dat" For Input As #1
   Print , "*** 學 生 成 績 表 ***"
   Print "姓 名", "國文", "英文", "平均"
   Do While Not EOF(1)
      Input #1, nam, chinese, english
      Print nam, chinese, english, (chinese + english) / 2
   Loop
   Close #1
End Sub

