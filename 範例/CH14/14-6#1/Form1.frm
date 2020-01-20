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
   CurrentX = 400: CurrentY = 400  '設定輸出起點的座標
   Print "(300,500)"               '輸出起點的座標
   DrawWidth = 2                   '設定線條寬度
   Line (300, 500)-(1000, 1200)    '畫直線
   Print "("; CurrentX; ","; CurrentY; ")" '輸出終點的座標
End Sub

