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
   Dim a                        '設定a為不定型變數
   Dim b As Variant             '設定b為不定型變數
   Dim c(5)                     '設定c為不定型變數陣列
   Dim d As String              '設定d為字串變數
   a = Array("松", "竹", "梅")  '正確
   b = Array("松", "竹", "梅")  '正確
   c = Array("松", "竹", "梅")  '錯誤:無法指定至陣列
   d = Array("松", "竹", "梅")  '錯誤:型態不符
End Sub

