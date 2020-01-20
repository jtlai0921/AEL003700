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
   Dim qty As Single, price As Single, money As Single
   qty = InputBox("請輸入購書數量", "購書數量")
   price = InputBox("請輸入單價", "單價")
   money = price * qty
   If qty >= 10 Then money = money * 0.8
   Print "購書金額="; money; "元"
End Sub

