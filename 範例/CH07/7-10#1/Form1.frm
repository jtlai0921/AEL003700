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
   password = InputBox$("請輸入密碼", "密碼檢查")
   If password = "1234" Then
      MsgBox "通過密碼檢查了!", vbOKOnly + vbExclamation, _
             "恭喜!"
   Else
      feedback = MsgBox("密碼錯誤!", vbYesNoCancel _
                 + vbCritical, "抱歉!")
      Select Case feedback
         Case vbYes
            MsgBox "你按了「是(Y)」鈕", 0, "是"
         Case vbNo
            MsgBox "你按了「否(N)」鈕", 0, "否"
         Case vbCancel
            MsgBox "你按了「取消」鈕", 0, "取消"
         Case Else
            MsgBox "不應有此情況"
      End Select
   End If
End Sub

