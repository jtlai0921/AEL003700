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
  Dim s As Integer, g As Integer, t As Integer
  Randomize
  s = Int(Rnd * 50) + 1
    Do
    t = t + 1
    g = InputBox("請輸入你所猜的數值(1~50)")
    Select Case g
      Case s
        MsgBox "恭喜!你猜對了! 總共猜了" + Str(t) + "次"
        Exit Do
      Case Is < s
        MsgBox "第" + Str(t) + "次猜的數值" + Str(g) + "太低了!請再猜一次!"
      Case Else
        MsgBox "第" + Str(t) + "次猜的數值" + Str(g) + "太高了!請再猜一次!"
    End Select
  Loop
End Sub

