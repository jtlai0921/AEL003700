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
  Dim x As Integer, y As Integer, sum As Integer, score As Integer
  Randomize
  For i = 1 To 10
    x = Int(Rnd * 90) + 10       '產生10~99之間的隨意整數
    y = Int(Rnd * 90) + 10       '產生10~99之間的隨意整數
    sum = InputBox("請輸入" + Str(x) + "+" + Str(y) + "=?")
    If sum = x + y Then          '答對的情況
      MsgBox "恭喜!你這題答對了!"
      score = score + 10
    Else                         '答錯的情況
      MsgBox "抱歉!你這題答錯了! 正確的答案是:" + Str(x + y)
    End If
  Next i
  MsgBox "測驗結束!你的得分是:" + Str(score)
End Sub

