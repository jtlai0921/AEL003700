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
   Dim p As Single, r As Single, t As Single  'p為本金,r為年利率,t為本利和
   Dim y As Integer       'y為年數
   p = InputBox("請輸入本金", "輸入本金")
   r = InputBox("請輸入年利率", "輸入年利率")
   y = 1
   Print "年數", "本利和"
   Print                 '空一列
   Do While t < 2 * p '當t<2*p就繼續執行迴圈,否則結束
      t = p * (1 + r / 100) ^ y  '計算到y年的本利和
      Print y, t                 '輸出年數與本利和
      y = y + 1                  '累加一年
   Loop
End Sub

