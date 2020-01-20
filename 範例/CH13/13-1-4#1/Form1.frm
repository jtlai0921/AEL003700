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
  x = InputBox("請輸入x值")
  Select Case Sgn(x)
    Case -1
      y = 3 * x * x + 2 * x + 5
    Case 0
      y = 0
    Case 1
      y = 3 * x * x + 4 * x + 5
  End Select
  Print "y="; y
End Sub
