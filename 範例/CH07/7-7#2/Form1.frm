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
  s = InputBox("請輸入分數")
  Select Case s
     Case Is < 60
        Print "不及格"
        Print "請多用功！"
     Case 60 To 89
        Print "及格"
        Print "恭喜！"
     Case Else
        Print "優等"
        Print "發獎狀一張！"
  End Select
End Sub

