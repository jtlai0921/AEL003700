VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4155
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Open "a:\score.dat" For Input As #1
   Print Tab(5); "*** 學 生 成 績 表 ***"
   Print "姓  名"; Tab(10); "國文"; Tab(16); "英文"; Tab(22); "平均"
   Do While Not EOF(1)
      Input #1, nam, chin, eng
      Print nam; Tab(10); chin; Tab(16); eng; Tab(22); _
                 (Val(chin) + Val(eng)) / 2
   Loop
   Close #1
End Sub

