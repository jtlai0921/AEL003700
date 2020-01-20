VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "儲蓄是美德"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4785
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
  Dim day_deposit As Double, total_deposit As Double  '當天存款與累積存款
  Dim days As Integer   '持續存款天數
  deposit = InputBox("請輸入第一天存款金額")
  days = InputBox("請輸入存款天數")
  Print "天數", "當天存款額", "累積存款額"  '由印表機輸出
  For i = 1 To days    '迴圈的次數從第一天到最後一天
    day_deposit = deposit * 2 ^ (i - 1) '每天存款為第一天存款的2^(i-1)倍
    total_deposit = total_deposit + day_deposit  '累積總存款金額
     Print i, Format(day_deposit, "#########,###") _
                     , Format(total_deposit, "#########,###") '由印表機輸出
  Next i
End Sub

