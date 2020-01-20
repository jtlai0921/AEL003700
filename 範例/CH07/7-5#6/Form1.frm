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
   Dim c As String
   c = InputBox$("請輸入一個字元")
   If c >= "A" And c <= "Z" Then
      Print "大寫字母"
   ElseIf c >= "a" And c <= "z" Then
      Print "小寫字母"
   ElseIf c >= "0" And c <= "9" Then
      Print "數字"
   Else
      Print "不是字母或數字"
   End If
End Sub

