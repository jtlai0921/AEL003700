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
   w = ScaleWidth  '取得表單可繪圖部分的寬度
   h = ScaleHeight '取得表單可繪圖部分的高度
   Do
      n = n + 1
      r = n * 100    '每圈的半徑增加100Twips
      '半徑超過寬度或高度的一半,就跳離迴圈
      If r > w / 2 Or r > h / 2 Then Exit Do
      Circle (w / 2, h / 2), r, QBColor(n)  '畫圓
  Loop
End Sub

