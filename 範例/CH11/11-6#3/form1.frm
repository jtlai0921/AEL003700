VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5265
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim w As Integer
   Dim week As Variant  '設定a為不定型變數
   Do
   w = InputBox("請輸入0~6的一個整數")
   If w >= 0 And w <= 6 Then Exit Do
   MsgBox "輸入數值超出範圍,請重新輸入"
   Loop
   week = Array("日", "一", "二", "三", "四", "五", "六")
   Print "星期"; week(w)
End Sub

