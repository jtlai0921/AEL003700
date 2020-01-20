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
   Dim nopass As Integer, pass As Integer  'pass為及格人數,nopass為不及格人數
   Dim score As Single, total As Single     'score為個人分數,total為全班總分
   Do
      score = InputBox("請輸入分數", "輸入分數")
      If score < 0 Then Exit Do   '如果分數為負值,就跳離迴圈
      If score <= 100 Then        '分數正常
         If score >= 60 Then pass = pass + 1 Else nopass = nopass + 1
         total = total + score    '累加全班總分
      Else                        '分數超過100,就輸出錯誤訊息
         MsgBox "輸入錯誤-分數超過100,請重新輸入", 48, "資料錯誤"
      End If
   Loop
   Rem 輸出結果
   Print "及格人數  ="; pass; "人"
   Print "不及格人數="; nopass; "人"
   Print "全班平均  ="; total / (pass + nopass); "分"
End Sub

