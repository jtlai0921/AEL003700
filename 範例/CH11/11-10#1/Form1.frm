VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1  '此敘述要安排在(一般)(宣告)中
Private Sub Form_Activate()
   Rem 設定動態陣列及輸入其內容
   Dim n() As String  '宣告為存放姓名的動態陣列
   Dim s() As Single  '宣告為存放分數的動態陣列
   rec = 0            '用來記錄輸入的筆數
   nam = InputBox$("請輸入第1個學生的姓名")
   Do While nam <> "end"    '輸入姓名不是end就繼續迴圈
      rec = rec + 1         '輸入筆數加1
      score = Val(InputBox("請輸入第" + Str(rec) + "個學生的分數"))
      ReDim Preserve n(rec)
      ReDim Preserve s(rec)
      n(rec) = nam
      s(rec) = score
      nam = InputBox$("請輸入第" + Str(rec + 1) + "個學生的姓名")
   Loop
   Rem 輸出結果
   Print "    學  生  成  績  表"
   Print "姓  名", "成  績"
   For i = 1 To rec
      Print n(i), s(i)
   Next i
End Sub

