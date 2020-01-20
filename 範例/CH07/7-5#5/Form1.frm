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
   score = InputBox("請輸入成績")
   If score < 60 Then
      Print "不及格"
      Print "請多用功！"
   ElseIf score < 90 Then
      Print "及格"
      Print "恭喜！"
   Else
      Print "優等"
      Print "發獎狀一張！"
   End If
End Sub

