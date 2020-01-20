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
   Dim x(0)      '宣告靜態陣列，固定有x(0)一個元素
   Dim y()       '宣告0個元素的陣列，事後可重新宣告
   ReDim y(6)    '正確，重新宣告為7個元素的陣列
   'ReDim x(5)    '錯誤，會出現「已宣告過陣列的維數」之訊息
   'ReDim y(10) As String '錯誤，會出現「不能改變陣列元素的資料型態」之訊息
End Sub

